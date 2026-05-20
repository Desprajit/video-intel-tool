import os
import json
import re
import tempfile
from datetime import datetime

from flask import Flask, request, jsonify, render_template, send_file
from flask_cors import CORS
from googleapiclient.discovery import build
from google import genai
from dotenv import load_dotenv

from pptx_builder import build_pptx

load_dotenv()

app = Flask(__name__)
CORS(app)

YOUTUBE_API_KEY = os.getenv("YOUTUBE_API_KEY")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash")

USE_GEMINI = os.getenv("USE_GEMINI", "0") == "1"

if not YOUTUBE_API_KEY:
    raise ValueError("Missing YOUTUBE_API_KEY in .env")
if not GEMINI_API_KEY:
    raise ValueError("Missing GEMINI_API_KEY in .env")


def get_youtube_service():
    return build("youtube", "v3", developerKey=YOUTUBE_API_KEY)


def safe_int(value, default=0):
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def search_channel(youtube, company_name):
    """Search for a company's YouTube channel."""
    response = youtube.search().list(
        q=f"{company_name} official",
        type="channel",
        part="snippet",
        maxResults=5
    ).execute()

    items = response.get("items", [])
    if not items:
        return None

    company_lower = company_name.lower()

    for item in items:
        title = item.get("snippet", {}).get("title", "").lower()
        if company_lower in title or any(word in title for word in company_lower.split()):
            return item.get("snippet", {}).get("channelId")

    return items[0].get("snippet", {}).get("channelId")


def get_channel_stats(youtube, channel_id):
    """Get channel statistics."""
    response = youtube.channels().list(
        part="snippet,statistics,contentDetails",
        id=channel_id
    ).execute()

    items = response.get("items", [])
    if not items:
        return {}

    item = items[0]
    stats = item.get("statistics", {})
    snippet = item.get("snippet", {})
    content_details = item.get("contentDetails", {})
    uploads_playlist = (
        content_details.get("relatedPlaylists", {})
        .get("uploads", "")
    )

    return {
        "channel_id": channel_id,
        "channel_name": snippet.get("title", "Unknown"),
        "description": snippet.get("description", "")[:300],
        "subscriber_count": safe_int(stats.get("subscriberCount")),
        "video_count": safe_int(stats.get("videoCount")),
        "view_count": safe_int(stats.get("viewCount")),
        "country": snippet.get("country", "N/A"),
        "published_at": snippet.get("publishedAt", "")[:10],
        "uploads_playlist": uploads_playlist
    }


def get_recent_videos(youtube, uploads_playlist_id, max_results=20):
    """Get recent videos from a channel."""
    if not uploads_playlist_id:
        return []

    playlist_response = youtube.playlistItems().list(
        part="snippet,contentDetails",
        playlistId=uploads_playlist_id,
        maxResults=max_results
    ).execute()

    playlist_items = playlist_response.get("items", [])
    video_ids = [
        item.get("contentDetails", {}).get("videoId")
        for item in playlist_items
        if item.get("contentDetails", {}).get("videoId")
    ]

    if not video_ids:
        return []

    stats_response = youtube.videos().list(
        part="snippet,statistics,contentDetails",
        id=",".join(video_ids)
    ).execute()

    videos = []
    for item in stats_response.get("items", []):
        stats = item.get("statistics", {})
        snippet = item.get("snippet", {})
        duration = item.get("contentDetails", {}).get("duration", "PT0S")

        views = safe_int(stats.get("viewCount"))
        likes = safe_int(stats.get("likeCount"))
        comments = safe_int(stats.get("commentCount"))

        engagement_rate = round(((likes + comments) / views) * 100, 2) if views > 0 else 0

        videos.append({
            "video_id": item.get("id", ""),
            "title": snippet.get("title", ""),
            "published_at": snippet.get("publishedAt", "")[:10],
            "views": views,
            "likes": likes,
            "comments": comments,
            "engagement_rate": engagement_rate,
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

    dates = [v["published_at"] for v in videos if v.get("published_at")]
    if len(dates) < 2:
        return "Unknown"

    try:
        parsed_dates = sorted(
            [datetime.strptime(d, "%Y-%m-%d") for d in dates],
            reverse=True
        )

        newest = parsed_dates[0]
        oldest = parsed_dates[-1]
        days_span = (newest - oldest).days

        if days_span == 0:
            return "Multiple per day"

        videos_per_week = round(len(parsed_dates) / (days_span / 7), 1)

        if videos_per_week >= 7:
            return "Daily"
        elif videos_per_week >= 3:
            return f"{videos_per_week} videos/week"
        elif videos_per_week >= 1:
            return f"{videos_per_week} videos/week"
        else:
            days_between = round(days_span / len(parsed_dates))
            return f"{days_between} days between posts"
    except Exception:
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


def extract_json_from_text(text):
    """Try to extract valid JSON from model output."""
    if not text:
        raise ValueError("Empty model response")

    text = text.strip()

    # Remove common markdown fences
    text = re.sub(r"^```json\s*", "", text)
    text = re.sub(r"^```\s*", "", text)
    text = re.sub(r"\s*```$", "", text)

    # First try direct parsing
    try:
        return json.loads(text)
    except Exception:
        pass

    # Then try to find the first JSON object in the text
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        return json.loads(match.group(0))

    raise ValueError("Could not parse JSON from model response")


def generate_ai_analysis(all_data, company_name):
    """Generate deep insights. Gemini is optional; fallback is always available."""

    def fallback_analysis(error_text="Fallback used"):
        fallback_rankings = [d["company"] for d in all_data if "error" not in d]
        return {
            "executive_summary": "This report analyzes competitor YouTube performance using real public video data including subscriber growth, engagement, posting consistency, and content strategy patterns.",
            "leader": all_data[0]["company"] if all_data else company_name,
            "leader_reason": "Fallback analysis used because Gemini was disabled or unavailable.",
            "content_themes": {},
            "content_gaps": ["Short-form content", "Behind the scenes", "Customer stories", "Live streams"],
            "posting_insight": "Consistent posting usually improves visibility and audience recall.",
            "engagement_insight": "Educational and practical videos often receive stronger engagement.",
            "recommendations": [
                {"title": "Increase posting frequency", "detail": "Publish at least 2 times per week to stay visible and build consistency."},
                {"title": "Add Shorts", "detail": "Use YouTube Shorts to improve discovery and attract new viewers."},
                {"title": "Publish case studies", "detail": "Customer stories build trust and often perform well with B2B audiences."},
                {"title": "Improve thumbnails", "detail": "Test thumbnail styles to improve click-through rate."},
                {"title": "Engage in comments", "detail": "Reply to viewers to strengthen community signals and engagement."}
            ],
            "company_scores": {},
            "rankings": fallback_rankings,
            "missing_formats": ["Shorts", "Live streams", "Webinars"]
        }

    if not USE_GEMINI:
        return fallback_analysis("Gemini disabled")

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

    prompt = f"""
You are a senior video marketing strategist.

Analyze this YouTube data for {company_name} and its competitors.

DATA:
{json.dumps(summary_data, indent=2)}

Return ONLY valid JSON with this structure:
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

Rules:
- Make 5 recommendations
- Score companies out of 10
- Keep it concise
- No markdown
"""

    client = genai.Client(api_key=GEMINI_API_KEY)

    try:
        response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,
            config={
                "response_mime_type": "application/json",
                "temperature": 0.2,
                "max_output_tokens": 2048,
            }
        )

        text = (response.text or "").strip()
        text = re.sub(r"^```json\s*", "", text)
        text = re.sub(r"^```\s*", "", text)
        text = re.sub(r"\s*```$", "", text)

        data = json.loads(text)

        # Fill missing keys safely
        data.setdefault("content_themes", {})
        data.setdefault("content_gaps", [])
        data.setdefault("recommendations", [])
        data.setdefault("company_scores", {})
        data.setdefault("rankings", [])
        data.setdefault("missing_formats", [])

        return data

    except Exception as e:
        print("Gemini error:", e)
        return fallback_analysis(str(e))


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze():
    data = request.json or {}
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
        ai_insights = {
            "executive_summary": f"AI analysis unavailable: {str(e)}",
            "leader": company,
            "leader_reason": "Fallback analysis used.",
            "content_themes": {},
            "content_gaps": [],
            "posting_insight": "",
            "engagement_insight": "",
            "recommendations": [],
            "company_scores": {},
            "rankings": [d["company"] for d in all_data if "error" not in d],
            "missing_formats": []
        }

    return jsonify({
        "companies": all_data,
        "insights": ai_insights,
        "main_company": company
    })


@app.route("/download", methods=["POST"])
def download():
    data = request.json or {}
    report_data = data.get("report_data")

    if not report_data:
        return jsonify({"error": "No report data"}), 400

    try:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
        tmp.close()

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