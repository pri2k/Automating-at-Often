import asyncio
import os
from dotenv import load_dotenv
from firecrawl import AsyncFirecrawlApp

load_dotenv()

FIRECRAWL_API_KEY = os.getenv("FIRECRAWL")
TARGET_URL = "https://www.skyscanner.co.in/transport/flights/del/dxb/250601/?adultsv2=1&cabinclass=economy&childrenv2=&inboundaltsenabled=false&outboundaltsenabled=false&ref=home&rtn=0&stops=!twoPlusStops"

async def crawl_page(url):
    app = AsyncFirecrawlApp(api_key=FIRECRAWL_API_KEY)
    response = await app.scrape_url(
        url=url,
        formats=['markdown', 'html', 'screenshot@fullPage'],
        # only_main_content=True
    )
    return response

async def main():
    print("🔍 Crawling Skyscanner page...")
    result = await crawl_page(TARGET_URL)

    # print("Using API key:", FIRECRAWL_API_KEY)

    
    if not result or 'html' not in result:
        print("❌ No HTML returned. Try with a different URL or ensure login is not required.")
        return
    
    html = result['html']
    with open("skyscanner.html", "w", encoding="utf-8") as f:
        f.write(html)
    
    print("✅ Saved HTML to skyscanner.html")

if __name__ == "__main__":
    asyncio.run(main())
