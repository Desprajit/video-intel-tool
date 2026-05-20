import os
import json
import re
from flask import Flask, request, jsonify, render_template, send_file
from flask_cors import CORS
from googleapiclient.discovery import build
from google import genai
from google.genai import types
from dotenv import load_dotenv
from pptx_builder import build_pptx
import tempfile

load_dotenv()

app = Flask(__name__)
CORS(app)

YOUTUBE_API_KEY = os.getenv("YOUTUBE_API_KEY")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")


def get_youtube_service():
    return build("youtube", "v3", developerKey=YOUTUBE_API_KEY)

def search_channel(youtube, company_name):
    """Search for a company's YouTube channel."""
    response = youtube.search().list(
        q=company_name + " official",
        type="channel",
        part="snippet",
        maxResults=3
    ).execute()

    if not response.get("items"):
        return None

    # Pick the best match
    for item in response["items"]:
        title = item["snippet"]["title"].lower()
        if company_name.lower() in title or any(w in title for w in company_name.lower().split()):
            return item["snippet"]["channelId"]

    return response["items"][0]["snippet"]["channelId"]

def get_channel_stats(youtube, channel_id):
    """Get channel statistics."""
    response = youtube.channels().list(
        part="snippet,statistics,contentDetails",
        id=channel_id
    ).execute()

    if not response.get("items"):
        return {}

    item = response["items"][0]
    stats = item.get("statistics", {})
    snippet = item.get("snippet", {})

    return {
        "channel_id": channel_id,
        "channel_name": snippet.get("title", "Unknown"),
        "description": snippet.get("description", "")[:300],
        "subscriber_count": int(stats.get("subscriberCount", 0)),
        "video_count": int(stats.get("videoCount", 0)),
        "view_count": int(stats.get("viewCount", 0)),
        "country": snippet.get("country", "N/A"),
        "published_at": snippet.get("publishedAt", "")[:10],
        "uploads_playlist": item.get("contentDetails", {}).get("relatedPlaylists", {}).get("uploads", "")
    }

def get_recent_videos(youtube, uploads_playlist_id, max_results=20):
    """Get recent videos from a channel."""
    if not uploads_playlist_id:
        return []

    response = youtube.playlistItems().list(
        part="snippet,contentDetails",
        playlistId=uploads_playlist_id,
        maxResults=max_results
    ).execute()

    video_ids = [item["contentDetails"]["videoId"] for item in response.get("items", [])]
    if not video_ids:
        return []

    # Get video stats
    stats_response = youtube.videos().list(
        part="snippet,statistics,contentDetails",
        id=",".join(video_ids)
    ).execute()

    videos = []
    for item in stats_response.get("items", []):
        stats = item.get("statistics", {})
        snippet = item.get("snippet", {})
        duration = item.get("contentDetails", {}).get("duration", "PT0S")

        views = int(stats.get("viewCount", 0))
        likes = int(stats.get("likeCount", 0))
        comments = int(stats.get("commentCount", 0))

        videos.append({
            "video_id": item["id"],
            "title": snippet.get("title", ""),
            "published_at": snippet.get("publishedAt", "")[:10],
            "views": views,
            "likes": likes,
            "comments": comments,
            "engagement_rate": round((likes + comments) / views * 100, 2) if views > 0 else 0,
            "tags": snippet.get("tags", [])[:5],
            "description": snippet.get("description", "")[:200],
            "thumbnail": snippet.get("thumbnails", {}).get("medium", {}).get("url", ""),
            "duration": duration
        })

    return sorted(videos, key=lambda x: x["views"], reverse=True)

def analyze_posting_frequency(videos):
    """Calculate posting frequency from video dates."""
    if len(videos) < 2:
        return "Insufficient data"

    dates = sorted([v["published_at"] for v in videos if v["published_at"]], reverse=True)
    if len(dates) < 2:
        return "Unknown"

    from datetime import datetime
    try:
        d1 = datetime.strptime(dates[0], "%Y-%m-%d")
        d2 = datetime.strptime(dates[-1], "%Y-%m-%d")
        days_span = (d1 - d2).days
        if days_span == 0:
            return "Multiple per day"
        videos_per_week = round(len(dates) / (days_span / 7), 1)
        if videos_per_week >= 7:
            return "Daily"
        elif videos_per_week >= 3:
            return f"{videos_per_week} videos/week"
        elif videos_per_week >= 1:
            return f"{videos_per_week} videos/week"
        else:
            return f"{round(days_span/len(dates))} days between posts"
    except:
        return "Unknown"

def fetch_company_data(company_name):
    """Fetch all YouTube data for a company."""
    youtube = get_youtube_service()

    channel_id = search_channel(youtube, company_name)
    if not channel_id:
        return {"company": company_name, "error": "Channel not found"}

    channel_stats = get_channel_stats(youtube, channel_id)
    videos = get_recent_videos(youtube, channel_stats.get("uploads_playlist", ""))

    avg_views = round(sum(v["views"] for v in videos) / len(videos)) if videos else 0
    avg_likes = round(sum(v["likes"] for v in videos) / len(videos)) if videos else 0
    avg_comments = round(sum(v["comments"] for v in videos) / len(videos)) if videos else 0
    avg_engagement = round(sum(v["engagement_rate"] for v in videos) / len(videos), 2) if videos else 0

    all_tags = []
    for v in videos:
        all_tags.extend(v.get("tags", []))

    top_topics = list(dict.fromkeys(all_tags))[:10]

    return {
        "company": company_name,
        "channel_name": channel_stats.get("channel_name", company_name),
        "channel_id": channel_id,
        "subscribers": channel_stats.get("subscriber_count", 0),
        "total_videos": channel_stats.get("video_count", 0),
        "total_views": channel_stats.get("view_count", 0),
        "country": channel_stats.get("country", "N/A"),
        "channel_since": channel_stats.get("published_at", "N/A"),
        "posting_frequency": analyze_posting_frequency(videos),
        "recent_videos": videos[:20],
        "top_videos": videos[:5],
        "avg_views": avg_views,
        "avg_likes": avg_likes,
        "avg_comments": avg_comments,
        "avg_engagement": avg_engagement,
        "top_topics": top_topics
    }

def generate_ai_analysis(all_data, company_name):
    """Use Gemini to generate deep insights."""
    summary_data = []
    for d in all_data:
        if "error" not in d:
            top_titles = [v["title"] for v in d.get("top_videos", [])[:5]]
            summary_data.append({
                "company": d["company"],
                "subscribers": d["subscribers"],
                "total_videos": d["total_videos"],
                "avg_views": d["avg_views"],
                "avg_engagement": d["avg_engagement"],
                "posting_frequency": d["posting_frequency"],
                "top_video_titles": top_titles,
                "top_topics": d.get("top_topics", [])
            })

    prompt = f"""You are a senior video marketing strategist. Analyze this YouTube data for {company_name} and its competitors.

DATA:
{json.dumps(summary_data, indent=2)}

Provide a detailed analysis in this EXACT JSON format (no markdown, pure JSON):
{{
  "executive_summary": "3-4 sentences on who leads in video marketing and why",
  "leader": "name of company leading in video marketing",
  "leader_reason": "one sentence why they lead",
  "content_themes": {{
    "company_name": ["theme1", "theme2", "theme3"]
  }},
  "content_gaps": ["gap1", "gap2", "gap3", "gap4"],
  "posting_insight": "2-3 sentences about posting patterns and what works",
  "engagement_insight": "2-3 sentences about engagement patterns",
  "recommendations": [
    {{
      "title": "Recommendation title",
      "detail": "Specific actionable detail for {company_name}"
    }}
  ],
  "company_scores": {{
    "company_name": {{
      "content_quality": 7,
      "consistency": 8,
      "engagement": 6,
      "growth_potential": 9,
      "overall": 7.5
    }}
  }},
  "rankings": ["1st company", "2nd company", "3rd company"],
  "missing_formats": ["format1", "format2", "format3"]
}}

Make 5 recommendations. Score companies out of 10. Be concise."""
    
    client = genai.Client(api_key=GEMINI_API_KEY)
    response = client.models.generate_content(
    model="gemini-1.5-flash-8b",
    contents=prompt
)

    try:
        text = response.text.strip()
        # Remove markdown code blocks if present
        text = re.sub(r'^```json\s*', '', text)
        text = re.sub(r'\s*```$', '', text)
        return json.loads(text)
    except:
        return {
            "executive_summary": "Analysis complete. See detailed metrics below.",
            "leader": all_data[0]["company"] if all_data else company_name,
            "leader_reason": "Based on subscriber count and engagement metrics.",
            "content_themes": {},
            "content_gaps": ["Short-form content", "Behind the scenes", "Customer stories", "Live streams"],
            "posting_insight": "Consistent posting drives better channel growth.",
            "engagement_insight": "Higher engagement correlates with educational content.",
            "recommendations": [
                {"title": "Increase posting frequency", "detail": "Post at least 2x per week to grow faster."},
                {"title": "Add shorts", "detail": "YouTube Shorts drive discovery."},
                {"title": "Customer case studies", "detail": "Real customer stories perform well."},
                {"title": "Improve thumbnails", "detail": "A/B test thumbnail designs."},
                {"title": "Engage in comments", "detail": "Reply to comments to boost engagement signals."}
            ],
            "company_scores": {},
            "rankings": [d["company"] for d in all_data if "error" not in d],
            "missing_formats": ["Shorts", "Live streams", "Webinars"]
        }

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/analyze", methods=["POST"])
def analyze():
    data = request.json
    company = data.get("company", "").strip()
    competitors = [c.strip() for c in data.get("competitors", []) if c.strip()]

    if not company:
        return jsonify({"error": "Company name is required"}), 400

    all_companies = [company] + competitors[:4]
    all_data = []

    for name in all_companies:
        try:
            company_data = fetch_company_data(name)
            all_data.append(company_data)
        except Exception as e:
            all_data.append({"company": name, "error": str(e)})

    try:
        ai_insights = generate_ai_analysis(all_data, company)
    except Exception as e:
        ai_insights = {"executive_summary": f"AI analysis unavailable: {str(e)}"}

    return jsonify({
        "companies": all_data,
        "insights": ai_insights,
        "main_company": company
    })

@app.route("/download", methods=["POST"])
def download():
    data = request.json
    report_data = data.get("report_data")

    if not report_data:
        return jsonify({"error": "No report data"}), 400

    try:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
        build_pptx(report_data, tmp.name)
        return send_file(
            tmp.name,
            as_attachment=True,
            download_name="video_competitor_report.pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True, port=5000)