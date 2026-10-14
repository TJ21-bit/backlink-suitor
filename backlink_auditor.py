#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════╗
║       BACKLINK AUDITOR v3.1 — Local AI Edition           ║
║  Playwright + Local AI heuristic. No external services.  ║
╚══════════════════════════════════════════════════════════╝

HOW TO RUN:
    python backlink_auditor.py "Tijeet singh.xlsx"

FIRST TIME SETUP (run once):
    pip install requests beautifulsoup4 openpyxl pandas playwright
    python -m playwright install chromium

CHECKS:
    1. Backlink Live or Dead          — with exact reason
    2. Target Keyword on page         — anchor, title, body, H1-H3
    3. Page links back to Project URL — checks <a>, data-attrs, JS, redirects
    4. AI Content Score (0-10)        — local heuristic (phrases + burstiness)
    5. DA / SS columns                — blank for now, ready for Moz later
"""

import argparse
import json
import os
import re
import sys
import time
from urllib.parse import urlparse

import requests
import pandas as pd
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Colors ─────────────────────────────────────────────────────────────────────
CLR_HEADER = "FF1565C0"
CLR_GREEN  = "FFB9F6CA"
CLR_RED    = "FFFFCDD2"
CLR_YELLOW = "FFFFF9C4"
CLR_BLUE   = "FFE3F2FD"
CLR_PURPLE = "FFF3E5F5"
CLR_WHITE  = "FFFFFFFF"


# ══════════════════════════════════════════════════════════════════════════════
# 1. PAGE FETCH
# ══════════════════════════════════════════════════════════════════════════════

def _scroll_page(page):
    """Quick 2-step scroll to trigger lazy-loaded content."""
    try:
        page.evaluate("""
            async () => {
                window.scrollTo(0, document.body.scrollHeight / 2);
                await new Promise(r => setTimeout(r, 300));
                window.scrollTo(0, document.body.scrollHeight);
                await new Promise(r => setTimeout(r, 200));
                window.scrollTo(0, 0);
            }
        """)
        page.wait_for_timeout(200)
    except Exception:
        pass


# ── Known short URL domains ──────────────────────────────────────────────────
SHORT_URL_DOMAINS = {
    "bit.ly", "tinyurl.com", "t.co", "t.ly", "goo.gl",
    "ow.ly", "is.gd", "cutt.ly", "rebrand.ly", "shorturl.at",
    "tiny.cc", "bl.ink", "buff.ly", "dlvr.it", "fur.ly",
    "mcaf.ee", "short.io", "soo.gd", "s.id", "v.gd",
    "clck.ru", "qr.ae", "adf.ly", "bit.do",
    # Image sharing short URLs
    "pin.it", "flic.kr", "img.ly",
}

# Paste/document sites — NOT shorteners, open directly
# These host actual content, not redirectors
PASTE_SITE_DOMAINS = {
    "justpaste.me", "justpaste.it", "pastebin.com", "paste.ee",
    "rentry.co", "privatebin.net", "hastebin.com", "dpaste.org",
}


def _is_short_url(url: str) -> bool:
    """Returns True if URL is a known short URL service (not a paste/content site)."""
    try:
        domain = urlparse(url).netloc.replace("www.", "").lower()
        if domain in PASTE_SITE_DOMAINS:
            return False   # paste sites host content, not redirectors
        return domain in SHORT_URL_DOMAINS
    except Exception:
        return False


def _resolve_short_url(url: str) -> str:
    """
    Resolve a short URL to its final destination.
    Tries multiple strategies:
    1. HEAD request following redirects (fastest)
    2. GET request with stream=True (catches more redirect types)
    3. Check response history for intermediate redirects
    Returns final URL, or original url if all strategies fail.
    """
    headers = {
        "User-Agent"     : "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Accept"         : "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
    }
    try:
        session = requests.Session()
        session.max_redirects = 15

        # Strategy 1: HEAD with redirects
        try:
            resp = session.head(url, headers=headers, timeout=10, allow_redirects=True)
            if resp.url and resp.url.rstrip("/") != url.rstrip("/"):
                return resp.url
            # Check redirect history
            if resp.history:
                last = resp.history[-1]
                loc = last.headers.get("Location", "")
                if loc and loc.startswith("http"):
                    return loc
        except Exception:
            pass

        # Strategy 2: GET with stream to avoid downloading full page
        try:
            resp = session.get(url, headers=headers, timeout=10,
                               allow_redirects=True, stream=True)
            resp.close()
            if resp.url and resp.url.rstrip("/") != url.rstrip("/"):
                return resp.url
            if resp.history:
                return resp.history[-1].headers.get("Location", url) or url
        except Exception:
            pass

    except Exception:
        pass

    return url


def _goto(page, url: str) -> dict:
    """
    Navigate to URL using a smart two-stage wait.
    If the URL is a known short URL, resolve it first via HTTP HEAD.
    Also fixes common URL issues (double slashes).

    Returns { live, code, error, html, final_url, was_short_link }
    """
    result = {
        "live"           : False,
        "code"           : None,
        "error"          : None,
        "html"           : "",
        "final_url"      : url,
        "was_short_link" : False,
    }

    # ── Fix common URL issues ─────────────────────────────────────────────
    # Fix double slashes after domain e.g. https://instaconnect.co//post/...
    from urllib.parse import urlsplit, urlunsplit
    try:
        parts = urlsplit(url)
        clean_path = re.sub(r'/{2,}', '/', parts.path)
        url = urlunsplit((parts.scheme, parts.netloc, clean_path, parts.query, parts.fragment))
        result["final_url"] = url
    except Exception:
        pass

    # ── Pre-resolve short URLs ────────────────────────────────────────────
    actual_url = url
    if _is_short_url(url):
        resolved = _resolve_short_url(url)
        if resolved and resolved != url:
            actual_url               = resolved
            result["final_url"]      = resolved
            result["was_short_link"] = True
            print(f"          🔗 Short URL resolved: {url[:40]} → {resolved[:50]}")
        else:
            print(f"          🔗 Short URL — trying browser directly: {url[:50]}")

    try:
        # Stage 1: Load the page with a generous timeout, but only wait for DOM ready
        resp = page.goto(actual_url, wait_until="domcontentloaded", timeout=45000)
        code = resp.status if resp else None
        result["code"] = code

        # Stage 2: Try networkidle but only wait 5s max
        # Many sites never reach true idle due to ads/analytics
        try:
            page.wait_for_load_state("networkidle", timeout=5000)
        except PlaywrightTimeout:
            page.wait_for_timeout(1500)

        # Capture final URL after all redirects
        try:
            playwright_final = page.url
            orig_domain  = urlparse(url).netloc.replace("www.", "").lower()
            final_domain = urlparse(playwright_final).netloc.replace("www.", "").lower()

            if not result["was_short_link"]:
                # Not already resolved — check if Playwright followed a redirect
                result["final_url"] = playwright_final
                if orig_domain and final_domain and orig_domain != final_domain:
                    result["was_short_link"] = True
            else:
                # Already resolved via HTTP — but update to Playwright's
                # final URL in case there were more hops after the short link
                if playwright_final and playwright_final != actual_url:
                    result["final_url"] = playwright_final
        except Exception:
            pass

        # Now extract content
        def grab_html():
            html = page.content()
            for frame in page.frames[1:]:
                try:
                    html += frame.content()
                except Exception:
                    pass
            return html

        def is_login_page(html: str, url: str) -> bool:
            """
            Detect if the page IS a login wall — meaning the main content
            is a login/signup form and the actual post is inaccessible.
            Does NOT flag pages that merely have a login button in their header.
            """
            html_lower = html.lower()
            url_lower  = url.lower()

            # Signal 1: URL itself redirected to a login page
            # Only flag if the URL changed to contain login/signin
            login_url_signals = ["/login", "/signin", "/sign-in",
                                  "/auth/login", "/account/login", "/users/sign_in"]
            if any(url_lower.endswith(s) or f"{s}?" in url_lower
                   for s in login_url_signals):
                return True

            from bs4 import BeautifulSoup as _BS
            soup = _BS(html, "html.parser")

            # Signal 2: Page title is ONLY about login/register — very specific
            title = (soup.title.string or "").strip().lower() if soup.title else ""
            login_only_titles = ["login", "sign in", "signin", "log in",
                                  "register", "create account", "join"]
            # Must be EXACTLY a login title, not "Login | Hasster" where content exists
            if title in login_only_titles:
                return True

            # Signal 3: The MAIN/ONLY visible content is a login form
            # Remove header, footer, nav — check if what remains is just a login form
            for tag in soup(["header", "footer", "nav", "script", "style", "aside"]):
                tag.decompose()

            main_text = soup.get_text(separator=" ", strip=True).lower()

            # If page has very little text AND has a password field → login wall
            has_password = (
                'type="password"' in html_lower or
                "type='password'" in html_lower
            )
            # A real content page will have lots of text
            # A login wall will have very little
            if has_password and len(main_text) < 500:
                return True

            return False

        if code == 200:
            result["live"] = True
            _scroll_page(page)
            html = grab_html()

            # Check if we landed on a login page instead of actual content
            if is_login_page(html, page.url):
                result["live"]  = None
                result["html"]  = ""
                result["error"] = "Login required — content hidden behind login wall (verify manually)"
            else:
                result["html"] = html

        elif code is None:
            # Page loaded but no HTTP response captured (client redirect)
            html = grab_html()
            if html and len(html) > 1000:
                result["live"] = True
                _scroll_page(page)
                result["html"] = grab_html()
            else:
                result["live"]  = None
                result["error"] = "No HTTP status — page may have redirected (verify manually)"

        elif code == 403:
            # Playwright got 403 — try requests library as fallback
            # Some sites block headless Chrome fingerprint but allow plain HTTP
            try:
                fallback = requests.get(
                    actual_url,
                    headers={
                        "User-Agent"     : "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
                        "Accept"         : "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                        "Accept-Language": "en-US,en;q=0.9",
                        "Accept-Encoding": "gzip, deflate, br",
                        "Referer"        : "https://www.google.com/",
                    },
                    timeout=15,
                    allow_redirects=True,
                )
                if fallback.status_code == 200 and len(fallback.text) > 500:
                    result["live"]  = True
                    result["html"]  = fallback.text
                    result["code"]  = 200
                    result["final_url"] = fallback.url
                else:
                    result["live"]  = None
                    result["error"] = "HTTP 403 — bot blocked (verify manually)"
            except Exception:
                result["live"]  = None
                result["error"] = "HTTP 403 — bot blocked (verify manually)"
        else:
            # Temporary overload codes — retry once after 5 seconds
            temporary_codes = {
                429: "Too many requests — server rate-limited us",
                500: "Server error — may be temporary",
                502: "Bad gateway — may be temporary",
                503: "Service unavailable — may be temporary",
                508: "Resource limit reached — server temporarily overloaded",
                509: "Bandwidth exceeded — server temporarily overloaded",
            }
            permanent_codes = {
                404: "Page not found",
                410: "Page permanently removed",
                400: "Bad request — URL may be malformed",
                401: "Login required",
            }

            if code in temporary_codes:
                # Wait 5 seconds and retry once
                import time as _time
                _time.sleep(5)
                try:
                    resp2 = page.goto(actual_url, wait_until="domcontentloaded", timeout=30000)
                    code2 = resp2.status if resp2 else None
                    if code2 == 200:
                        result["live"] = True
                        _scroll_page(page)
                        result["html"] = page.content()
                        result["code"] = 200
                    else:
                        result["live"]  = None
                        result["error"] = f"HTTP {code} — {temporary_codes[code]} (verify manually)"
                except Exception:
                    result["live"]  = None
                    result["error"] = f"HTTP {code} — {temporary_codes[code]} (verify manually)"
            elif code in permanent_codes:
                result["error"] = f"HTTP {code} — {permanent_codes[code]}"
            else:
                result["error"] = f"HTTP {code} — unexpected response"

    except PlaywrightTimeout:
        # Initial DOM load failed even with 45s — try ONE more time with the
        # bare minimum wait condition: 'commit' (just first byte received)
        try:
            resp = page.goto(actual_url, wait_until="commit", timeout=20000)
            page.wait_for_timeout(4000)
            html = page.content()
            for frame in page.frames[1:]:
                try:
                    html += frame.content()
                except Exception:
                    pass

            code = resp.status if resp else None
            result["code"] = code

            # Capture final URL after redirects in fallback path too
            try:
                final_url = page.url
                result["final_url"] = final_url
                orig_domain  = urlparse(url).netloc.replace("www.", "").lower()
                final_domain = urlparse(final_url).netloc.replace("www.", "").lower()
                if orig_domain and final_domain and orig_domain != final_domain:
                    result["was_short_link"] = True
            except Exception:
                pass

            if html and len(html) > 1500:
                # Got useful content even though page didn't fully load
                if code == 200 or code is None:
                    result["live"]  = True
                    result["html"]  = html
                    result["error"] = None
                elif code == 403:
                    result["live"]  = None
                    result["error"] = "HTTP 403 — bot blocked (verify manually)"
                else:
                    result["error"] = f"HTTP {code} — slow load, partial content only"
            else:
                result["error"] = "Timeout (45s+20s) — site too slow or bot-blocking"
        except Exception:
            result["error"] = "Timeout (45s) — site too slow or bot-blocking"

    except Exception as e:
        err = str(e)
        if "ERR_NAME_NOT_RESOLVED" in err:
            result["error"] = "Domain not found — URL may be wrong"
        elif "ERR_CONNECTION_REFUSED" in err:
            result["error"] = "Connection refused"
        elif "ERR_CONNECTION_TIMED_OUT" in err or "ERR_TIMED_OUT" in err:
            result["error"] = "Connection timed out — server unreachable"
        elif "ERR_CONNECTION_CLOSED" in err:
            # Server closed connection — often means temporary overload (508/503)
            # Try once more with requests library as fallback
            try:
                import time as _time
                _time.sleep(3)
                fallback = requests.get(
                    actual_url,
                    headers={
                        "User-Agent"     : "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
                        "Accept"         : "text/html,application/xhtml+xml,*/*;q=0.8",
                        "Accept-Language": "en-US,en;q=0.9",
                        "Referer"        : "https://www.google.com/",
                    },
                    timeout=15,
                    allow_redirects=True,
                )
                if fallback.status_code == 200 and len(fallback.text) > 500:
                    result["live"]  = True
                    result["html"]  = fallback.text
                    result["code"]  = 200
                    result["final_url"] = fallback.url
                elif fallback.status_code in (508, 509, 503, 429):
                    result["live"]  = None
                    result["error"] = f"HTTP {fallback.status_code} — server temporarily overloaded (verify manually)"
                else:
                    result["live"]  = None
                    result["error"] = f"Connection closed — server may be temporarily overloaded (verify manually)"
            except Exception:
                result["live"]  = None
                result["error"] = "Connection closed — server temporarily overloaded (verify manually)"
        elif "ERR_SSL" in err:
            result["error"] = "SSL certificate error"
        elif "interrupted by another" in err.lower() or "navigation" in err.lower():
            result["error"] = "Navigation interrupted — site redirects too much"
        else:
            clean = err.replace("Page.goto: ", "").split(" at http")[0]
            result["error"] = clean[:80] if clean else "Unknown error"

    return result


# ══════════════════════════════════════════════════════════════════════════════
# 2. KEYWORD CHECK
# ══════════════════════════════════════════════════════════════════════════════

# ── PDF and image hosting domains — keyword check not applicable ─────────────
PDF_HOSTING_DOMAINS = {
    "dropbox.com", "4shared.com", "slideshare.net", "scribd.com",
    "issuu.com", "calameo.com", "docdroid.net", "mediafire.com",
    "docs.google.com", "drive.google.com",
    # Flipbook / PDF viewer platforms
    "heyzine.com", "files.fm", "smallpdf.com", "flipsnack.com",
    "yumpu.com", "pubhtml5.com", "joomag.com", "flipbuilder.com",
    # File sharing with PDF preview
    "jumpshare.com", "filestack.com", "box.com", "onedrive.live.com",
}
IMAGE_HOSTING_DOMAINS = {
    "behance.net", "flickr.com", "pinterest.com", "instagram.com",
    "500px.com", "deviantart.com", "imgur.com", "photobucket.com",
    "unsplash.com", "pexels.com",
}

def _is_pdf_url(url: str) -> bool:
    domain = urlparse(url).netloc.replace("www.", "").lower()
    return domain in PDF_HOSTING_DOMAINS or url.lower().endswith(".pdf")

def _is_image_url(url: str) -> bool:
    domain = urlparse(url).netloc.replace("www.", "").lower()
    return domain in IMAGE_HOSTING_DOMAINS


def check_keyword(html: str, keyword: str) -> tuple:
    """
    Returns (found: bool, where: str|None).
    Checks exact keyword AND common variations:
    - "skiltravel" also checks "skil travel", "skil-travel"
    - "seo services delhi" also checks "seoservicesdelhi", "seo-services-delhi"
    """
    kw   = keyword.strip().lower()
    soup = BeautifulSoup(html, "html.parser")

    # Build all keyword variants to check
    variants = [kw]
    if " " in kw:
        # Multi-word: also try no-space and hyphenated
        variants.append(kw.replace(" ", ""))
        variants.append(kw.replace(" ", "-"))
    else:
        # Single compound word: try with spaces and hyphens at split points
        # e.g. "skiltravel" → "skil travel", "skil-travel"
        for i in range(3, len(kw) - 2):
            variants.append(kw[:i] + " " + kw[i:])
            variants.append(kw[:i] + "-" + kw[i:])

    # Remove duplicates
    variants = list(dict.fromkeys(variants))

    # Build text targets once
    title_text = (soup.title.string or "").lower() if soup.title else ""
    meta_tag   = soup.find("meta", attrs={"name": "description"})
    meta_text  = (meta_tag.get("content") or "").lower() if meta_tag else ""
    body_text  = soup.get_text(separator=" ", strip=True).lower()
    anchors    = [(a.get_text(strip=True) or "").lower() for a in soup.find_all("a")]
    headings   = [(h.get_text(strip=True).lower(), h.name.upper()) for h in soup.find_all(["h1","h2","h3"])]

    for variant in variants:
        if variant in title_text:
            return True, "Page Title"
        if variant in meta_text:
            return True, "Meta Description"
        for anchor in anchors:
            if variant in anchor:
                return True, "Anchor Text"
        for heading_text, tag_name in headings:
            if variant in heading_text:
                return True, tag_name
        if variant in body_text:
            return True, "Page Body"

    return False, None


# ══════════════════════════════════════════════════════════════════════════════
# 3. PROJECT URL LINK CHECK
# ══════════════════════════════════════════════════════════════════════════════

def check_project_url_linked(html: str, project_url: str) -> dict:
    """
    Check if backlink page links to the submitted project URL.
    Handles: direct links, tracking redirects, data-* attrs,
    JS-embedded URLs, short URLs in href AND in plain text.
    Returns { linked, dofollow, link_text, match_type }
    """
    if not project_url or not html:
        return {"linked": False, "dofollow": None, "link_text": "—", "match_type": None}

    from urllib.parse import unquote

    soup           = BeautifulSoup(html, "html.parser")
    project_parsed = urlparse(project_url)
    project_domain = project_parsed.netloc.replace("www.", "").lower()
    project_path   = project_parsed.path.rstrip("/").lower()
    project_full   = project_url.rstrip("/").lower()

    if not project_domain:
        return {"linked": False, "dofollow": None, "link_text": "—", "match_type": None}

    # ── STAGE 1: Standard <a href=...> tags ─────────────────────────────────
    for a in soup.find_all("a", href=True):
        raw_href = a.get("href") or ""
        href     = raw_href.lower()
        decoded  = unquote(href)

        href_domain = urlparse(href).netloc.replace("www.", "").lower()
        href_path   = urlparse(href).path.rstrip("/").lower()

        match_type = None
        if href_domain == project_domain:
            match_type = "Direct link"
        elif project_full in decoded:
            match_type = "Via redirect"
        elif project_domain in decoded:
            match_type = "Domain match"
        elif project_path and len(project_path) > 3 and href_path == project_path:
            match_type = "Path match"

        # ── Check if href is a short URL that resolves to project URL ─────
        if not match_type and _is_short_url(raw_href):
            resolved = _resolve_short_url(raw_href)
            if resolved and project_domain in resolved.lower():
                match_type = "Short link (resolved)"

        if match_type:
            rel         = a.get("rel") or []
            is_nofollow = "nofollow" in " ".join(rel).lower()
            link_text   = a.get_text(strip=True) or "—"
            return {
                "linked"    : True,
                "dofollow"  : not is_nofollow,
                "link_text" : link_text[:60],
                "match_type": match_type,
            }

    # ── STAGE 2: Other attributes that often hold URLs ──────────────────────
    for tag in soup.find_all(True):
        for attr in ("data-href", "data-url", "data-link", "data-original-url",
                     "data-target", "src", "data-src"):
            val = tag.get(attr)
            if val:
                val_lower   = val.lower()
                val_decoded = unquote(val_lower)
                if (project_domain in val_decoded
                        or project_full in val_decoded):
                    text = tag.get_text(strip=True) or tag.get("title", "") or "—"
                    return {
                        "linked"    : True,
                        "dofollow"  : True,
                        "link_text" : text[:60],
                        "match_type": f"Via {attr}",
                    }

    # ── STAGE 3: Short URLs in plain page text ───────────────────────────────
    # Handles: "Read More:- https://shorturl.at/LpjC4" where the short URL
    # resolves to the project URL. Scans entire page text for any short URL
    # pattern and resolves each one.
    page_text = soup.get_text(separator=" ", strip=True)
    # Find all URLs in the page text
    url_pattern = re.compile(
        r'https?://[^\s<>"\')\]]+',
        re.IGNORECASE
    )
    all_urls_in_text = url_pattern.findall(page_text)
    for found_url in all_urls_in_text:
        found_clean = found_url.rstrip(".,;:!?)")  # strip trailing punctuation
        if _is_short_url(found_clean):
            resolved = _resolve_short_url(found_clean)
            if resolved and project_domain in resolved.lower():
                return {
                    "linked"    : True,
                    "dofollow"  : True,
                    "link_text" : f"Short URL in text → {found_clean}",
                    "match_type": "Short URL in text (resolved)",
                }

    # ── STAGE 4: Raw HTML scan — full URL anywhere in source ─────────────────
    html_lower   = html.lower()
    html_decoded = unquote(html_lower)

    if project_full in html_decoded or project_url.lower() in html_decoded:
        return {
            "linked"    : True,
            "dofollow"  : True,
            "link_text" : "Found in page source",
            "match_type": "JS / embedded",
        }

    # ── STAGE 5: Domain-only match in raw HTML ───────────────────────────────
    if project_domain in html_decoded:
        idx     = html_decoded.find(project_domain)
        context = html_decoded[max(0, idx-200):idx+200]
        link_signals = ["href", "src", "url", "link", "source", "http", "www"]
        if any(sig in context for sig in link_signals):
            return {
                "linked"    : True,
                "dofollow"  : True,
                "link_text" : "Domain found in page source",
                "match_type": "Domain in source",
            }

    # ── STAGE 6: Plain text mention (not a real backlink) ────────────────────
    if project_domain in page_text.lower():
        return {
            "linked"    : False,
            "dofollow"  : None,
            "link_text" : "URL mentioned but NOT clickable",
            "match_type": "text-only",
        }

    return {"linked": False, "dofollow": None, "link_text": "—", "match_type": None}

    return {"linked": False, "dofollow": None, "link_text": "—", "match_type": None}


# ══════════════════════════════════════════════════════════════════════════════
# 4. AI CONTENT SCORE — Local heuristic detector (no external service needed)
# ══════════════════════════════════════════════════════════════════════════════

# Common AI-writing patterns
AI_PHRASES = [
    "delve into", "delving into", "navigate the", "navigating the",
    "in today's digital", "in the realm of", "in the world of",
    "it's important to note", "it is important to note",
    "it's worth noting", "it is worth noting",
    "in conclusion", "to conclude", "in summary",
    "furthermore", "moreover", "additionally",
    "first and foremost", "last but not least",
    "a testament to", "tapestry of", "landscape of",
    "play a pivotal role", "play a crucial role", "play a key role",
    "in essence", "in a nutshell",
    "embark on", "embarking on",
    "comprehensive guide", "comprehensive overview",
    "let's explore", "let us explore",
    "the ever-evolving", "ever-changing",
    "shed light on", "sheds light on",
    "underscores the importance", "highlights the importance",
    "by leveraging", "leveraging the",
    "intricacies of", "complexities of",
    "stands out as", "stands as a",
    "in the digital age", "in this digital era",
    "transformative", "revolutionary",
    "seamlessly", "seamless integration",
    "cutting-edge", "state-of-the-art",
]

AI_TRANSITION_WORDS = {
    "furthermore", "moreover", "additionally", "consequently",
    "therefore", "thus", "hence", "subsequently",
    "nevertheless", "nonetheless", "however",
    "specifically", "particularly", "notably",
}


def _extract_page_text(html: str) -> str:
    """Strip boilerplate from HTML and return clean body text."""
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "nav", "footer", "header",
                     "aside", "form", "noscript"]):
        tag.decompose()
    text = soup.get_text(separator=" ", strip=True)
    return " ".join(text.split())


def get_ai_score(html: str, ai_page=None) -> dict:
    """
    Local AI-detection heuristic. No external service, no API key.
    Looks for AI-writing fingerprints:
      1. AI catchphrases (delve into, tapestry, navigate, etc.)
      2. Overuse of transition words
      3. Sentence length uniformity (low burstiness)
      4. Average sentence complexity

    Returns { score_10, ai_pct, label, error }
    Score: 0% AI = 10/10 (human), 100% AI = 0/10 (likely AI)
    """
    result = {"score_10": "—", "ai_pct": None, "label": "—", "error": None}

    if not html:
        result["error"] = "No content"
        return result

    text = _extract_page_text(html)

    if len(text) < 100:
        result["label"]    = "Too short to score"
        result["score_10"] = "—"
        return result

    text_lower = text.lower()
    words      = text_lower.split()
    word_count = len(words)

    if word_count < 50:
        result["label"]    = "Too short to score"
        result["score_10"] = "—"
        return result

    # ─ Signal 1: AI catchphrases per 1000 words ─────────────────────────────
    phrase_hits = sum(text_lower.count(p) for p in AI_PHRASES)
    phrases_per_1k = (phrase_hits / word_count) * 1000

    # ─ Signal 2: Transition word density ────────────────────────────────────
    transition_hits = sum(1 for w in words if w.strip(".,;:!?") in AI_TRANSITION_WORDS)
    transitions_per_1k = (transition_hits / word_count) * 1000

    # ─ Signal 3: Sentence length variance (burstiness) ──────────────────────
    sentences = re.split(r"[.!?]+", text)
    sentence_lengths = [len(s.split()) for s in sentences if len(s.split()) >= 3]

    if len(sentence_lengths) >= 5:
        mean_len = sum(sentence_lengths) / len(sentence_lengths)
        variance = sum((x - mean_len) ** 2 for x in sentence_lengths) / len(sentence_lengths)
        std_dev  = variance ** 0.5
        # Low std deviation = AI-like uniformity
        # Humans: std dev usually 6-15. AI: usually 2-6.
        burstiness_score = max(0, min(100, (8 - std_dev) * 12))  # higher = more AI-like
    else:
        burstiness_score = 30   # neutral

    # ─ Combine signals into AI probability % ────────────────────────────────
    # Tunable weights:
    phrase_signal     = min(100, phrases_per_1k * 18)        # 5 phrases/1k → 90
    transition_signal = min(100, transitions_per_1k * 1.5)   # 30/1k → 45

    ai_pct = (
        phrase_signal     * 0.45 +
        burstiness_score  * 0.35 +
        transition_signal * 0.20
    )
    ai_pct = round(max(0.0, min(100.0, ai_pct)), 1)

    score_10 = round(10 - (ai_pct / 10), 1)
    score_10 = max(0.0, min(10.0, score_10))

    if ai_pct <= 15:
        label = "Human Written"
    elif ai_pct <= 35:
        label = "Mostly Human"
    elif ai_pct <= 60:
        label = "Mixed Content"
    elif ai_pct <= 80:
        label = "Mostly AI"
    else:
        label = "AI Generated"

    result["ai_pct"]   = ai_pct
    result["score_10"] = score_10
    result["label"]    = label
    return result


# ══════════════════════════════════════════════════════════════════════════════
# 5. AUDIT ONE ROW
# ══════════════════════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════════════════════
# 5b. KNOWN LOW-QUALITY BOOKMARKING PLATFORMS
# ══════════════════════════════════════════════════════════════════════════════
# These platforms use Sngine/PHPDolphin scripts and hide outbound links from
# non-logged-in visitors. Our scraper can't detect the link reliably, but the
# link likely IS there for logged-in users — so we mark them for manual review
# instead of falsely flagging as "Not Linked".

LOW_QUALITY_PLATFORMS = {
    "fri3nd.me", "followingbook.com", "culturesbook.com",
    "topkif.nvinio.com", "we2chat.net", "antspride.com",
    "bytevidsocial.com", "instaconnect.co", "lebanonhub.app",
    "druzefaces.com", "yoursocial.it.com", "zirvede.net",
    "pikumil.com", "get2knowgod.com",
    # Community platforms requiring login
    "mn.co", "network-1194656.mn.co", "janjaonline.mn.co",
    # Messaging/social apps — profile content hidden without login
    "imo.net", "whoosmind.com", "ekcochat.com",
    "stayhomenetwork.com", "purekonect.com",
}


def is_low_quality_platform(url: str) -> bool:
    """Returns True if URL belongs to a known low-quality or login-walled platform."""
    try:
        domain = urlparse(url).netloc.replace("www.", "").lower()
        # Exact match
        if domain in LOW_QUALITY_PLATFORMS:
            return True
        # Subdomain match — e.g. janjaonline.mn.co → matches mn.co
        parts = domain.split(".")
        if len(parts) >= 2:
            parent = ".".join(parts[-2:])
            if parent in LOW_QUALITY_PLATFORMS:
                return True
        return False
    except Exception:
        return False


def _deep_retry_project_url(page, project_url: str) -> dict:
    """
    Last-resort project URL detection for stubborn community-bookmarking sites
    where the link loads via AJAX after page render. Performs:
      1. Extra scroll + wait for AJAX
      2. Direct DOM query for ALL <a> elements (post-render)
      3. Check every href, data-* attribute, and the entire rendered page text
    Returns same shape as check_project_url_linked().
    """
    from urllib.parse import unquote
    try:
        # Force more JS execution time
        page.wait_for_timeout(2500)

        # Aggressive scroll — some sites only load post body on scroll-into-view
        page.evaluate("""
            async () => {
                const sleep = ms => new Promise(r => setTimeout(r, ms));
                window.scrollTo(0, 0); await sleep(400);
                window.scrollTo(0, document.body.scrollHeight / 2); await sleep(500);
                window.scrollTo(0, document.body.scrollHeight); await sleep(500);
                window.scrollTo(0, document.body.scrollHeight / 2); await sleep(400);
                window.scrollTo(0, 0); await sleep(400);
            }
        """)
        page.wait_for_timeout(1500)

        project_parsed = urlparse(project_url)
        project_domain = project_parsed.netloc.replace("www.", "").lower()
        project_full   = project_url.rstrip("/").lower()

        # Direct DOM query — gets links as they exist AFTER all JS has run
        anchors = page.evaluate("""
            () => Array.from(document.querySelectorAll('a')).map(a => ({
                href: a.href || '',
                text: (a.innerText || a.textContent || '').trim().slice(0, 80),
                rel : a.getAttribute('rel') || ''
            }))
        """)

        for a in anchors:
            href    = (a.get("href") or "").lower()
            decoded = unquote(href)
            href_domain = urlparse(href).netloc.replace("www.", "").lower()

            if (href_domain == project_domain
                    or project_full in decoded
                    or project_domain in decoded):
                is_nofollow = "nofollow" in (a.get("rel") or "").lower()
                return {
                    "linked"    : True,
                    "dofollow"  : not is_nofollow,
                    "link_text" : (a.get("text") or "—")[:60],
                    "match_type": "Deep scan",
                }

        # Also check all data-* attributes via JS
        data_attr_hits = page.evaluate(f"""
            () => {{
                const target = '{project_domain}';
                const all = document.querySelectorAll('*');
                const hits = [];
                for (const el of all) {{
                    for (const attr of el.attributes) {{
                        const v = (attr.value || '').toLowerCase();
                        if (v.includes(target)) {{
                            hits.push({{tag: el.tagName, attr: attr.name, val: v.slice(0,100), text: (el.innerText||'').slice(0,40)}});
                            break;
                        }}
                    }}
                    if (hits.length > 3) break;
                }}
                return hits;
            }}
        """)

        if data_attr_hits:
            hit = data_attr_hits[0]
            return {
                "linked"    : True,
                "dofollow"  : True,
                "link_text" : hit.get("text", "—") or "—",
                "match_type": f"Deep ({hit.get('attr','attr')})",
            }

        # Final fallback — rendered page source (after JS) contains URL anywhere
        rendered = page.content().lower()
        if project_domain in unquote(rendered):
            return {
                "linked"    : True,
                "dofollow"  : True,
                "link_text" : "Found in rendered page",
                "match_type": "Deep / rendered",
            }

    except Exception:
        pass

    return {"linked": False, "dofollow": None, "link_text": "—", "match_type": None}


def audit_row(row: dict, page, ai_page) -> dict:
    live_url      = str(row.get("Backlink Live Link",    "") or "").strip()
    keyword       = str(row.get("Target Keyword",        "") or "").strip()
    project_url   = str(row.get("Submitted Project Url", "") or "").strip()
    activity_type = str(row.get("Backlink Activity Type","") or "").strip()

    flags, reasons = [], []

    # ── On-site Blog special handling ────────────────────────────────────────
    # For On-site Blog the URL IS the client's own website — no external
    # backlink check needed. Just verify page is live and keyword is present.
    is_onsite_blog = activity_type.lower() == "on-site blog"

    # ── Defaults ─────────────────────────────────────────────────────────────
    link_status  = "⛔ No URL provided"
    kw_display   = "⛔ N/A"
    proj_display = "⛔ N/A (on-site blog)" if is_onsite_blog else "⛔ N/A"
    ai_display   = "⛔ N/A"
    link_is_bot  = False
    link_is_dead = False

    if not live_url:
        flags.append("No URL")
        reasons.append("No backlink URL provided")

    else:
        # ── Step 1: Fetch backlink page ───────────────────────────────────
        fetch = _goto(page, live_url)

        link_is_live  = fetch["live"] is True
        link_is_bot   = fetch["live"] is None
        link_is_dead  = fetch["live"] is False
        link_is_login = link_is_bot and "Login required" in (fetch.get("error") or "")

        if link_is_live:
            if fetch.get("was_short_link"):
                final_domain = urlparse(fetch["final_url"]).netloc.replace("www.", "")
                link_status  = f'✅ Live (Short link → {final_domain})'
            else:
                link_status = "✅ Live"
        elif link_is_login:
            link_status = "⚠️ Login required — content hidden behind login wall"
        elif link_is_bot:
            link_status = "⚠️ Verify manually — bot blocked (HTTP 403)"
        else:
            link_status = f'❌ {fetch["error"]}' if fetch["error"] else "❌ Dead"

        if link_is_dead:
            flags.append("Dead Link")
            reasons.append(f'Dead link — {fetch["error"]}')
        if link_is_login:
            flags.append("Login Wall")
            reasons.append("Page shows login screen — backlink content not publicly visible, no SEO value")

        # ── Step 2: Keyword check ─────────────────────────────────────────
        if link_is_dead:
            kw_display = "⛔ N/A (link dead)"
        elif link_is_login:
            kw_display = "⛔ N/A (login required)"
        elif link_is_bot:
            kw_display = "⚠️ Verify manually (bot blocked)"
        elif not keyword:
            kw_display = "⛔ N/A (no keyword)"
        elif _is_pdf_url(fetch.get("final_url", live_url)):
            kw_display = "⛔ N/A (PDF — cannot check keyword in document)"
        elif _is_image_url(fetch.get("final_url", live_url)):
            kw_display = "⚠️ Image platform — keyword in image title/desc only"
            kw_found, kw_where = check_keyword(fetch["html"], keyword)
            if kw_found:
                kw_display = f"✅ Yes ({kw_where})"
        else:
            # First try on raw HTML
            kw_found, kw_where = check_keyword(fetch["html"], keyword)

            # If not found — try again using Playwright's fully rendered text
            # This catches JS-rendered content that BeautifulSoup misses
            if not kw_found:
                try:
                    rendered_text = page.inner_text("body").lower()
                    if keyword.lower().strip() in rendered_text:
                        kw_found = True
                        kw_where = "Page Body (JS rendered)"
                except Exception:
                    pass

            if kw_found:
                kw_display = f"✅ Yes ({kw_where})"
            else:
                if is_low_quality_platform(fetch.get("final_url", live_url)):
                    kw_display = "⚠️ Verify manually (content may be behind login)"
                else:
                    kw_display = "❌ Not found"
                    flags.append("Keyword Missing")
                    reasons.append(f'"{keyword}" not found on page')

        # ── Step 3: Project URL backlink check ────────────────────────────
        if is_onsite_blog:
            proj_display = "⛔ N/A (on-site blog — internal page)"
        elif link_is_dead:
            proj_display = "⛔ N/A (link dead)"
        elif link_is_login:
            proj_display = "⛔ N/A (login required — content not public)"
        elif link_is_bot:
            proj_display = "⚠️ Verify manually (bot blocked)"
        elif not project_url:
            proj_display = "⛔ N/A (no project URL)"
        else:
            proj = check_project_url_linked(fetch["html"], project_url)

            # If not found on first try, do a deep DOM scan (post-AJAX)
            # This catches community-bookmarking sites that load the link via JS
            if not proj["linked"] and proj.get("match_type") != "text-only":
                deep = _deep_retry_project_url(page, project_url)
                if deep["linked"]:
                    proj = deep

            if proj["linked"]:
                follow_tag   = "Dofollow ✅" if proj["dofollow"] else "Nofollow ⚠️"
                match_tag    = proj.get("match_type") or "Linked"
                proj_display = f'✅ {match_tag} ({follow_tag}) — "{proj["link_text"]}"'
                if not proj["dofollow"]:
                    reasons.append("Link is Nofollow — less SEO value")
            elif proj.get("match_type") == "text-only":
                proj_display = "⚠️ URL mentioned in text but NOT a clickable link"
                flags.append("Project URL Not a Backlink")
                reasons.append("Project URL is only plain text — no actual <a> link")
            elif is_low_quality_platform(fetch.get("final_url", live_url)):
                # Known Sngine/PHPDolphin clone — link likely hidden from non-logged-in users
                proj_display = "⚠️ Known low-quality platform — verify manually"
                reasons.append(
                    "Backlink is on a Sngine/PHPDolphin-style social platform — "
                    "outbound link may only be visible to logged-in users, "
                    "please verify manually"
                )
            else:
                proj_display = "❌ Not linked"
                flags.append("Project URL Not Linked")
                reasons.append(f"Page does not link back to project URL")

        # ── Step 4: AI content score ──────────────────────────────────────
        if link_is_dead:
            ai_display = "⛔ N/A (link dead)"
        elif link_is_login:
            ai_display = "⛔ N/A (login required)"
        elif link_is_bot:
            ai_display = "⚠️ Verify manually (bot blocked)"
        else:
            ai = get_ai_score(fetch["html"], ai_page)
            if ai["error"]:
                ai_display = f'⚠️ Unavailable ({ai["error"]})'
            else:
                ai_display = f'{ai["score_10"]}/10 — {ai["label"]} ({ai["ai_pct"]}% AI)'
                if ai["ai_pct"] is not None and ai["ai_pct"] > 80:
                    flags.append("High AI Content")
                    reasons.append(
                        f'Content is {ai["ai_pct"]}% AI-written (score: {ai["score_10"]}/10)'
                    )

    # ── Final audit status ────────────────────────────────────────────────────
    # Use final URL (after redirects) for low-quality check
    final_check_url = ""
    try:
        final_check_url = fetch.get("final_url", live_url) if live_url else ""
    except Exception:
        final_check_url = live_url or ""

    is_low_quality = final_check_url and is_low_quality_platform(final_check_url)

    if not live_url:
        status = "🚩 Flagged"
    elif link_is_login:
        status = "🚩 Flagged — Login Wall (zero SEO value)"
    elif link_is_bot and not flags:
        status = "⚠️ Manual Check"
    elif link_is_bot and flags:
        status = "🚩 Flagged + ⚠️ Manual Check"
    elif is_low_quality and not flags:
        status = "⚠️ Manual Check"
    elif is_low_quality and flags:
        status = "🚩 Flagged + ⚠️ Manual Check"
    else:
        status = "✅ Good" if not flags else "🚩 Flagged"

    # Capture final URL for output (defaults to original if no fetch happened)
    final_url_output = ""
    if live_url:
        try:
            # 'fetch' only exists if we got past the URL check above
            final_url_output = fetch.get("final_url", live_url) if fetch.get("was_short_link") else ""
        except Exception:
            final_url_output = ""

    return {
        "Link Status"         : link_status,
        "Final URL"           : final_url_output,
        "Keyword Found?"      : kw_display,
        "Project URL Linked?" : proj_display,
        "AI Content Score"    : ai_display,
        "Audit Result"        : status,
        "Flag Reason"         : "; ".join(reasons) if reasons else "All checks passed ✓",
    }


# ══════════════════════════════════════════════════════════════════════════════
# 6. EXCEL STYLING
# ══════════════════════════════════════════════════════════════════════════════

def apply_styling(path: str):
    wb = load_workbook(path)
    ws = wb.active
    thin   = Side(style="thin", color="FFCCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    hdr_fill = PatternFill("solid", start_color=CLR_HEADER)
    hdr_font = Font(bold=True, color=CLR_WHITE, size=10, name="Arial")
    for cell in ws[1]:
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.border    = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32

    col_map = {cell.value: cell.column for cell in ws[1]}
    fills   = {
        "good": PatternFill("solid", start_color=CLR_GREEN),
        "bad" : PatternFill("solid", start_color=CLR_RED),
        "warn": PatternFill("solid", start_color=CLR_YELLOW),
        "info": PatternFill("solid", start_color=CLR_BLUE),
    }

    def getcell(row, name):
        idx = col_map.get(name)
        return row[idx - 1] if idx else None

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=True, horizontal="left")
            cell.border    = border
            cell.font      = Font(name="Arial", size=9)

        # Audit Result
        ar = getcell(row, "Audit Result")
        if ar:
            val = str(ar.value or "")
            ar.fill = (fills["good"] if "Good" in val
                       else fills["warn"] if "Manual" in val
                       else fills["bad"])
            ar.font = Font(bold=True, name="Arial", size=9)
            ar.alignment = Alignment(horizontal="center", vertical="center")

        # Link Status
        ls = getcell(row, "Link Status")
        if ls:
            val = str(ls.value or "")
            ls.fill = (fills["good"] if "✅" in val
                       else fills["warn"] if "⚠️" in val
                       else fills["bad"])
            ls.alignment = Alignment(horizontal="center", vertical="center")

        # Keyword Found
        kf = getcell(row, "Keyword Found?")
        if kf:
            val = str(kf.value or "")
            kf.fill = (fills["good"] if "✅" in val
                       else fills["warn"] if "⚠️" in val
                       else fills["bad"])
            kf.alignment = Alignment(horizontal="center", vertical="center")

        # Project URL Linked
        pu = getcell(row, "Project URL Linked?")
        if pu:
            val = str(pu.value or "")
            pu.fill = (fills["good"] if "✅" in val and "Nofollow" not in val
                       else fills["warn"] if "Nofollow" in val or "⚠️" in val
                       else fills["bad"])
            pu.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        # AI Content Score
        ai = getcell(row, "AI Content Score")
        if ai:
            val = str(ai.value or "")
            if "Human Written" in val or "Mostly Human" in val:
                ai.fill = fills["good"]
            elif "Mixed" in val or "⚠️" in val:
                ai.fill = fills["warn"]
            elif "Mostly AI" in val or "AI Generated" in val:
                ai.fill = fills["bad"]
            else:
                ai.fill = fills["info"]
            ai.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        # DA / SS — center, no color until Moz is added
        for col_name in ("DA", "SS"):
            c = getcell(row, col_name)
            if c:
                c.alignment = Alignment(horizontal="center", vertical="center")

    widths = {
        "Backlink Create Date"  : 18,
        "Website"               : 16,
        "Domain Country"        : 14,
        "Employee Name"         : 18,
        "Backlink Activity Type": 20,
        "Referring Domain"      : 32,
        "DA"                    : 10,
        "SS"                    : 10,
        "Backlink Live Link"    : 52,
        "Target Keyword"        : 28,
        "Submitted Project Url" : 45,
        "Status"                : 10,
        "Link Status"           : 30,
        "Final URL"             : 50,
        "Keyword Found?"        : 24,
        "Project URL Linked?"   : 45,
        "AI Content Score"      : 32,
        "Audit Result"          : 16,
        "Flag Reason"           : 58,
        "Email"                 : 26,
        "User  Name"            : 14,
        "Password"              : 14,
    }
    for name, width in widths.items():
        if name in col_map:
            ws.column_dimensions[get_column_letter(col_map[name])].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


# ══════════════════════════════════════════════════════════════════════════════
# 7. MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description="Backlink Auditor v3 — AI Content Edition")
    parser.add_argument("input",    help="Excel file, e.g. 'Tijeet singh.xlsx'")
    parser.add_argument("--output", default="", help="Output filename (optional)")
    parser.add_argument(
        "--filter-methods",
        default="",
        help="Comma-separated list of methods to audit.",
    )
    # Meta fields — filled from UI if not in sheet
    parser.add_argument("--meta-project",    default="", help="Project/Website name")
    parser.add_argument("--meta-project-id", default="", help="Project ID")
    parser.add_argument("--meta-manager",    default="", help="Project Manager name")
    parser.add_argument("--meta-teamlead",   default="", help="Team Lead name")
    args = parser.parse_args()

    output = args.output or args.input.replace(".xlsx", "_AUDITED.xlsx")

    # Parse filter list (empty = audit everything)
    filter_methods = []
    if args.filter_methods:
        filter_methods = [m.strip() for m in args.filter_methods.split(",") if m.strip()]

    print("\n" + "═" * 64)
    print("   BACKLINK AUDITOR v3 — AI Content Edition")
    print("═" * 64)
    print(f"   Input   : {args.input}")
    print(f"   Output  : {output}")
    print(f"   Engine  : Headless Chrome + Local AI Heuristic Scoring")
    print(f"   DA / SS : — (add --moz-key later to enable)")
    if filter_methods:
        print(f"   Filter  : Auditing only — {', '.join(filter_methods)}")
    print("═" * 64 + "\n")

    # Auto-detect sheet name — try "Backlinks" first, then first available sheet
    xl        = pd.ExcelFile(args.input)
    all_sheets = xl.sheet_names
    if "Backlinks" in all_sheets:
        sheet_name = "Backlinks"
    elif "Worksheet" in all_sheets:
        sheet_name = "Worksheet"
    else:
        sheet_name = all_sheets[0]
    print(f"   Sheet   : {sheet_name}")

    df    = pd.read_excel(args.input, sheet_name=sheet_name, dtype=str).fillna("")
    total = len(df)
    print(f"   📋  {total} backlinks loaded.\n")

    # Find the activity type column for filtering
    activity_col = None
    for col in df.columns:
        if "activity type" in col.lower() or "method" in col.lower():
            activity_col = col
            break

    # Check if DA/SS already have values in the sheet (e.g. filled by team manually)
    da_prefilled = "DA" in df.columns and df["DA"].str.strip().replace("", pd.NA).notna().any()
    ss_prefilled = "SS" in df.columns and df["SS"].str.strip().replace("", pd.NA).notna().any()

    if da_prefilled or ss_prefilled:
        print(f"   DA/SS   : ✅ Using existing values from sheet")
    else:
        print(f"   DA/SS   : — (not in sheet, will show blank)")

    # Drop only old AUTO audit columns — preserve original DA/SS if already filled
    drop_cols = [
        "DA (Auto)", "SS % (Auto)",
        "Link Status", "Final URL", "Keyword Found?",
        "Project URL Linked?", "AI Content Score",
        "Audit Result", "Flag Reason"
    ]
    # Only drop DA/SS if they were blank (not pre-filled)
    if not da_prefilled:
        drop_cols.append("DA")
    if not ss_prefilled:
        drop_cols.append("SS")

    df = df.drop(columns=[c for c in drop_cols if c in df.columns])

    # Add new tracking columns if not already present
    for col, default in [
        ("Project ID",      ""),
        ("Project Manager", ""),
        ("Team Lead",       ""),
    ]:
        if col not in df.columns:
            df.insert(df.columns.get_loc("Employee Name") + 1 if "Employee Name" in df.columns else len(df.columns), col, default)

    # Fill from CLI meta args if column is blank and arg was provided
    meta_map = {
        "Project ID"     : args.meta_project_id,
        "Project Manager": args.meta_manager,
        "Team Lead"      : args.meta_teamlead,
        "Website"        : args.meta_project,
    }
    for col, val in meta_map.items():
        if val and col in df.columns:
            # Only fill blank cells
            df[col] = df[col].apply(lambda x: val if str(x).strip() == "" else x)
        elif val and col == "Website" and col not in df.columns:
            df[col] = val

    results     = [None] * total
    skipped_count = 0
    WORKERS     = 3   # number of parallel browser tabs

    # Pre-build rows list with index for parallel processing
    rows_list = [(i, idx, row) for i, (idx, row) in enumerate(df.iterrows(), 1)]

    # Separate skipped and to-audit rows
    to_audit = []
    for i, idx, row in rows_list:
        row_method = str(row.get(activity_col, "")).strip() if activity_col else ""
        if filter_methods and row_method and row_method not in filter_methods:
            skipped_count += 1
            print(f"   [{i:>3}/{total}]  ⏭  Skipped — {row_method}")
            results[i-1] = {
                "Link Status"         : f"⏸ Skipped — {row_method} not selected",
                "Final URL"           : "",
                "Keyword Found?"      : "⏸ Not audited",
                "Project URL Linked?" : "⏸ Not audited",
                "AI Content Score"    : "⏸ Not audited",
                "Audit Result"        : "⏸ Skipped",
                "Flag Reason"         : f"Activity type '{row_method}' not selected for this audit run",
            }
        else:
            to_audit.append((i, row))

    print(f"   🔍 Auditing {len(to_audit)} rows with {WORKERS} parallel workers...\n")

    import threading
    import queue

    lock         = threading.Lock()
    work_queue   = queue.Queue()
    for item in to_audit:
        work_queue.put(item)

    def worker(worker_id: int):
        """Each worker gets its own browser page and processes rows from the queue."""
        with sync_playwright() as pw:
            browser = pw.chromium.launch(
                headless=True,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-features=IsolateOrigins,site-per-process",
                ],
            )
            context = browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                ),
                viewport={"width": 1280, "height": 800},
                locale="en-US",
                extra_http_headers={
                    "Accept"                 : "text/html,application/xhtml+xml,*/*;q=0.8",
                    "Accept-Language"        : "en-US,en;q=0.9",
                    "Accept-Encoding"        : "gzip, deflate, br",
                    "DNT"                    : "1",
                    "Upgrade-Insecure-Requests": "1",
                },
            )
            context.add_init_script("""
                Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});
                Object.defineProperty(navigator, 'languages', {get: () => ['en-US','en']});
                window.chrome = { runtime: {} };
            """)
            page = context.new_page()
            page.route(
                "**/*.{png,jpg,jpeg,gif,svg,woff,woff2,ttf,mp4,webm,ico}",
                lambda r: r.abort()
            )

            while True:
                try:
                    i, row = work_queue.get_nowait()
                except queue.Empty:
                    break

                ref = str(row.get("Referring Domain", "")).strip()
                kw  = str(row.get("Target Keyword",   "")).strip()

                with lock:
                    print(f"   [{i:>3}/{total}]  W{worker_id}  {ref[:38]:<38}  kw: {kw[:18]}")

                result = audit_row(row.to_dict(), page, None)
                results[i-1] = result

                icon = ("✅" if "Good"    in result["Audit Result"] else
                        "⚠️"  if "Manual"  in result["Audit Result"] else "🚩")
                with lock:
                    print(
                        f"          → {icon}  "
                        f"Live:{result['Link Status'][:6]}  "
                        f"KW:{result['Keyword Found?'][:6]}  "
                        f"Proj:{result['Project URL Linked?'][:6]}"
                    )

                work_queue.task_done()
                time.sleep(0.2)

            browser.close()

    # Launch workers in parallel threads
    threads = []
    for wid in range(1, WORKERS + 1):
        t = threading.Thread(target=worker, args=(wid,), daemon=True)
        t.start()
        threads.append(t)
        time.sleep(0.5)   # stagger startup slightly

    for t in threads:
        t.join()

    # Fill any None slots (shouldn't happen but just in case)
    results = [r if r is not None else {
        "Link Status": "⚠️ Worker error — verify manually",
        "Final URL": "", "Keyword Found?": "⚠️ Error",
        "Project URL Linked?": "⚠️ Error", "AI Content Score": "—",
        "Audit Result": "⚠️ Manual Check", "Flag Reason": "Worker processing error",
    } for r in results]

    df = pd.concat([df, pd.DataFrame(results)], axis=1)
    df.to_excel(output, index=False, sheet_name=sheet_name)
    apply_styling(output)

    good    = sum(1 for r in results if r and "Good"    in r.get("Audit Result",""))
    manual  = sum(1 for r in results if r and "Manual"  in r.get("Audit Result",""))
    flagged = sum(1 for r in results if r and "Flagged" in r.get("Audit Result",""))
    skipped = sum(1 for r in results if r and "Skipped" in r.get("Audit Result",""))

    print("\n" + "═" * 64)
    print("   AUDIT COMPLETE")
    print("═" * 64)
    print(f"   Total           : {total}")
    print(f"   ✅ Good         : {good}")
    print(f"   ⚠️  Manual Check : {manual}")
    print(f"   🚩 Flagged      : {flagged}")
    if skipped:
        print(f"   ⏸  Skipped      : {skipped}")
    print(f"   📁 Output       : {output}")
    print("═" * 64 + "\n")


if __name__ == "__main__":
    main()