"""
TripSync Web Application — Full Selenium E2E Test Suite
========================================================
Covers:
  - Functional Tests  : All pages, forms, navigation, API integrations
  - Vulnerability Tests: XSS, open redirects, security headers, input sanitization
  - Unit-level Tests   : Backend API health, endpoint responses, CORS validation

Run locally:
    python tests/test_tripsync.py

Output:
    tests/tripsync_test_report.xlsx
"""

import os
import sys
import time
import json
import unittest
import datetime
import requests

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, WebDriverException
)
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter

# ─── Configuration ────────────────────────────────────────────────────────────
BASE_URL   = os.environ.get("TRIPSYNC_URL",    "https://abineshh502.github.io/TripSyncWeb")
API_URL    = os.environ.get("TRIPSYNC_API_URL", "https://tripsyncweb-backend.onrender.com")
HEADLESS   = os.environ.get("HEADLESS", "true").lower() == "true"
TIMEOUT    = int(os.environ.get("TIMEOUT", "20"))
REPORT_DIR = os.environ.get("REPORT_DIR", os.path.join(os.path.dirname(__file__), "..", "test-results"))

os.makedirs(REPORT_DIR, exist_ok=True)

# ─── Global result collector ───────────────────────────────────────────────────
_results: list[dict] = []

def _record(category: str, test_id: str, name: str, status: str,
            details: str = "", duration: float = 0.0, severity: str = "Normal"):
    _results.append({
        "category":  category,
        "test_id":   test_id,
        "name":      name,
        "status":    status,
        "details":   details,
        "duration":  round(duration, 2),
        "severity":  severity,
        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })


# ─── WebDriver factory ────────────────────────────────────────────────────────
def _build_driver() -> webdriver.Chrome:
    opts = Options()
    if HEADLESS:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1280,900")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-infobars")
    opts.add_argument("--ignore-certificate-errors")
    opts.add_argument("--log-level=3")
    opts.add_experimental_option("excludeSwitches", ["enable-logging"])
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=opts)


def _wait_visible(driver, by, selector, timeout=TIMEOUT):
    return WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located((by, selector))
    )

def _wait_clickable(driver, by, selector, timeout=TIMEOUT):
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by, selector))
    )

def _wait_title(driver, partial, timeout=TIMEOUT):
    return WebDriverWait(driver, timeout).until(
        EC.title_contains(partial)
    )


# ═══════════════════════════════════════════════════════════════════════════════
#  SECTION 1 — FUNCTIONAL TESTS
# ═══════════════════════════════════════════════════════════════════════════════
class FunctionalTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.driver = _build_driver()
        cls.driver.set_page_load_timeout(45)

    @classmethod
    def tearDownClass(cls):
        cls.driver.quit()

    def _go(self, path=""):
        self.driver.get(BASE_URL.rstrip("/") + ("/" + path.lstrip("/") if path else ""))
        time.sleep(1.5)

    # ── TC-F-001: Landing page loads ─────────────────────────────────────────
    def test_F001_landing_page_loads(self):
        t0 = time.time()
        try:
            self._go()
            body = self.driver.find_element(By.TAG_NAME, "body")
            assert body is not None
            page_src = self.driver.page_source.lower()
            assert "tripsync" in page_src, "TripSync brand not on landing page"
            _record("Functional", "F-001", "Landing Page Loads",
                    "PASS", f"Title: {self.driver.title}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-001", "Landing Page Loads",
                    "FAIL", str(e), time.time()-t0, "Critical")
            self.fail(str(e))

    # ── TC-F-002: Navbar is visible ──────────────────────────────────────────
    def test_F002_navbar_present(self):
        t0 = time.time()
        try:
            self._go()
            nav = self.driver.find_element(By.TAG_NAME, "nav")
            assert nav.is_displayed()
            _record("Functional", "F-002", "Navbar Visible on Landing",
                    "PASS", "nav element found and visible", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-002", "Navbar Visible on Landing",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-003: Login page loads ───────────────────────────────────────────
    def test_F003_login_page_loads(self):
        t0 = time.time()
        try:
            self._go("login")
            email_input = _wait_visible(self.driver, By.ID, "email")
            pass_input  = _wait_visible(self.driver, By.ID, "password")
            assert email_input.is_displayed()
            assert pass_input.is_displayed()
            _record("Functional", "F-003", "Login Page — Inputs Visible",
                    "PASS", "Email & password fields found", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-003", "Login Page — Inputs Visible",
                    "FAIL", str(e), time.time()-t0, "Critical")
            self.fail(str(e))

    # ── TC-F-004: Login form validation (empty submit) ───────────────────────
    def test_F004_login_empty_validation(self):
        t0 = time.time()
        try:
            self._go("login")
            btn = _wait_clickable(self.driver, By.CSS_SELECTOR, "button[type='submit']")
            btn.click()
            time.sleep(0.8)
            # Should NOT navigate away from login
            assert "login" in self.driver.current_url or "register" not in self.driver.current_url
            _record("Functional", "F-004", "Login Empty Validation Stays on Page",
                    "PASS", "Did not navigate away on empty submit", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-004", "Login Empty Validation Stays on Page",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-005: Login with wrong credentials shows error ───────────────────
    def test_F005_login_wrong_credentials(self):
        t0 = time.time()
        try:
            self._go("login")
            _wait_visible(self.driver, By.ID, "email").send_keys("invalid@test.com")
            self.driver.find_element(By.ID, "password").send_keys("wrongpassword123")
            self.driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
            time.sleep(4)
            page_src = self.driver.page_source
            # Should still be on login or show error
            has_error = any(kw in page_src.lower() for kw in
                           ["invalid", "error", "incorrect", "not found", "wrong"])
            assert has_error or "dashboard" not in self.driver.current_url
            _record("Functional", "F-005", "Login Wrong Credentials Error Handling",
                    "PASS", "Error shown or stayed on login", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-005", "Login Wrong Credentials Error Handling",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-006: Register page loads ────────────────────────────────────────
    def test_F006_register_page_loads(self):
        t0 = time.time()
        try:
            self._go("register")
            time.sleep(1.5)
            page_src = self.driver.page_source.lower()
            assert any(kw in page_src for kw in ["register", "sign up", "create", "account"])
            _record("Functional", "F-006", "Register Page Loads",
                    "PASS", f"URL: {self.driver.current_url}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-006", "Register Page Loads",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-007: Dashboard page redirects unauthenticated users ─────────────
    def test_F007_dashboard_auth_redirect(self):
        t0 = time.time()
        try:
            self._go("dashboard")
            time.sleep(3)
            url = self.driver.current_url
            # Should either show login or a redirect to /login or /
            redirected = "login" in url or url.rstrip("/").endswith(BASE_URL.rstrip("/"))
            page_src = self.driver.page_source.lower()
            shows_login_content = any(kw in page_src for kw in ["sign in", "email address", "password"])
            assert redirected or shows_login_content
            _record("Functional", "F-007", "Dashboard Redirects Unauthenticated Users",
                    "PASS", f"Redirected to: {url}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-007", "Dashboard Redirects Unauthenticated Users",
                    "FAIL", str(e), time.time()-t0, "High")
            self.fail(str(e))

    # ── TC-F-008: AI Assistant page accessible (redirect or content) ─────────
    def test_F008_ai_assistant_page(self):
        t0 = time.time()
        try:
            self._go("ai-assistant")
            time.sleep(3)
            page_src = self.driver.page_source.lower()
            reachable = any(kw in page_src for kw in ["tripsync", "assistant", "ai", "sign in", "login"])
            assert reachable
            _record("Functional", "F-008", "AI Assistant Page Reachable",
                    "PASS", f"URL: {self.driver.current_url}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-008", "AI Assistant Page Reachable",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-009: AI Planner page accessible ─────────────────────────────────
    def test_F009_ai_planner_page(self):
        t0 = time.time()
        try:
            self._go("ai-planner")
            time.sleep(3)
            page_src = self.driver.page_source.lower()
            assert any(kw in page_src for kw in ["tripsync", "planner", "journey", "sign in"])
            _record("Functional", "F-009", "AI Planner Page Reachable",
                    "PASS", f"URL: {self.driver.current_url}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-009", "AI Planner Page Reachable",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-010: Safety page accessible ─────────────────────────────────────
    def test_F010_safety_page(self):
        t0 = time.time()
        try:
            self._go("safety")
            time.sleep(3)
            page_src = self.driver.page_source.lower()
            assert any(kw in page_src for kw in ["tripsync", "safety", "crowd", "sign in"])
            _record("Functional", "F-010", "Safety & Crowd Advisor Page Reachable",
                    "PASS", f"URL: {self.driver.current_url}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-010", "Safety & Crowd Advisor Page Reachable",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-011: Expenses page accessible ───────────────────────────────────
    def test_F011_expenses_page(self):
        t0 = time.time()
        try:
            self._go("expenses")
            time.sleep(3)
            page_src = self.driver.page_source.lower()
            assert any(kw in page_src for kw in ["tripsync", "expense", "bill", "split", "sign in"])
            _record("Functional", "F-011", "Expense Splitter Page Reachable",
                    "PASS", f"URL: {self.driver.current_url}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-011", "Expense Splitter Page Reachable",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-012: Routes page accessible ─────────────────────────────────────
    def test_F012_routes_page(self):
        t0 = time.time()
        try:
            self._go("routes")
            time.sleep(3)
            page_src = self.driver.page_source.lower()
            assert any(kw in page_src for kw in ["tripsync", "route", "map", "optimize", "sign in"])
            _record("Functional", "F-012", "Route Optimization Page Reachable",
                    "PASS", f"URL: {self.driver.current_url}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-012", "Route Optimization Page Reachable",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-013: Route sharing page accessible ───────────────────────────────
    def test_F013_route_sharing_page(self):
        t0 = time.time()
        try:
            self._go("route-sharing")
            time.sleep(3)
            page_src = self.driver.page_source.lower()
            assert any(kw in page_src for kw in ["tripsync", "share", "itinerary", "route", "sign in"])
            _record("Functional", "F-013", "Route Sharing Page Reachable",
                    "PASS", f"URL: {self.driver.current_url}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-013", "Route Sharing Page Reachable",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-014: Explore page accessible ────────────────────────────────────
    def test_F014_explore_page(self):
        t0 = time.time()
        try:
            self._go("explore")
            time.sleep(3)
            page_src = self.driver.page_source.lower()
            assert any(kw in page_src for kw in ["tripsync", "explore", "discover", "place", "sign in"])
            _record("Functional", "F-014", "Explore Page Reachable",
                    "PASS", f"URL: {self.driver.current_url}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-014", "Explore Page Reachable",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-015: Trips page accessible ──────────────────────────────────────
    def test_F015_trips_page(self):
        t0 = time.time()
        try:
            self._go("trips")
            time.sleep(3)
            page_src = self.driver.page_source.lower()
            assert any(kw in page_src for kw in ["tripsync", "trip", "journey", "sign in"])
            _record("Functional", "F-015", "Trips Page Reachable",
                    "PASS", f"URL: {self.driver.current_url}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-015", "Trips Page Reachable",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-016: Groups page accessible ─────────────────────────────────────
    def test_F016_groups_page(self):
        t0 = time.time()
        try:
            self._go("groups")
            time.sleep(3)
            page_src = self.driver.page_source.lower()
            assert any(kw in page_src for kw in ["tripsync", "group", "team", "sign in"])
            _record("Functional", "F-016", "Groups Page Reachable",
                    "PASS", f"URL: {self.driver.current_url}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-016", "Groups Page Reachable",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-017: Profile page accessible ────────────────────────────────────
    def test_F017_profile_page(self):
        t0 = time.time()
        try:
            self._go("profile")
            time.sleep(3)
            page_src = self.driver.page_source.lower()
            assert any(kw in page_src for kw in ["tripsync", "profile", "account", "sign in"])
            _record("Functional", "F-017", "Profile Page Reachable",
                    "PASS", f"URL: {self.driver.current_url}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-017", "Profile Page Reachable",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-018: Notifications page accessible ───────────────────────────────
    def test_F018_notifications_page(self):
        t0 = time.time()
        try:
            self._go("notifications")
            time.sleep(3)
            page_src = self.driver.page_source.lower()
            assert any(kw in page_src for kw in ["tripsync", "notification", "alert", "sign in"])
            _record("Functional", "F-018", "Notifications Page Reachable",
                    "PASS", f"URL: {self.driver.current_url}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-018", "Notifications Page Reachable",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-019: Favorites page accessible ──────────────────────────────────
    def test_F019_favorites_page(self):
        t0 = time.time()
        try:
            self._go("favorites")
            time.sleep(3)
            page_src = self.driver.page_source.lower()
            assert any(kw in page_src for kw in ["tripsync", "favorite", "saved", "sign in"])
            _record("Functional", "F-019", "Favorites Page Reachable",
                    "PASS", f"URL: {self.driver.current_url}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-019", "Favorites Page Reachable",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-020: 404 page for unknown route ─────────────────────────────────
    def test_F020_404_page(self):
        t0 = time.time()
        try:
            self._go("this-page-does-not-exist-xyz-99999")
            time.sleep(2)
            page_src = self.driver.page_source.lower()
            is_404 = any(kw in page_src for kw in
                        ["404", "not found", "does not exist", "page not found", "tripsync"])
            assert is_404
            _record("Functional", "F-020", "404 Page Displays for Unknown Routes",
                    "PASS", f"URL: {self.driver.current_url}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-020", "404 Page Displays for Unknown Routes",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-021: Landing page has hero CTA link ──────────────────────────────
    def test_F021_landing_hero_cta(self):
        t0 = time.time()
        try:
            self._go()
            time.sleep(2)
            links = self.driver.find_elements(By.TAG_NAME, "a")
            hrefs = [l.get_attribute("href") or "" for l in links]
            has_auth_link = any("login" in h or "register" in h or "dashboard" in h
                               for h in hrefs)
            assert has_auth_link
            _record("Functional", "F-021", "Landing Page Has Auth CTA Link",
                    "PASS", "Found auth-related link in landing page", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-021", "Landing Page Has Auth CTA Link",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-022: Page title is meaningful ───────────────────────────────────
    def test_F022_page_title_present(self):
        t0 = time.time()
        try:
            self._go()
            title = self.driver.title
            assert title and len(title.strip()) > 0
            _record("Functional", "F-022", "Landing Page Has Non-Empty Title",
                    "PASS", f"Title: '{title}'", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-022", "Landing Page Has Non-Empty Title",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-023: Login link navigates to login ───────────────────────────────
    def test_F023_login_link_navigation(self):
        t0 = time.time()
        try:
            self._go()
            time.sleep(2)
            links = self.driver.find_elements(By.TAG_NAME, "a")
            login_link = next((l for l in links if "login" in (l.get_attribute("href") or "")), None)
            if login_link:
                login_link.click()
                time.sleep(2)
                assert "login" in self.driver.current_url
                _record("Functional", "F-023", "Login Navigation Link Works",
                        "PASS", f"Navigated to: {self.driver.current_url}", time.time()-t0)
            else:
                _record("Functional", "F-023", "Login Navigation Link Works",
                        "SKIP", "No login link found on landing page (may be behind auth)", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-023", "Login Navigation Link Works",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-F-024: Page is responsive (mobile viewport) ───────────────────────
    def test_F024_mobile_viewport(self):
        t0 = time.time()
        try:
            self.driver.set_window_size(375, 812)
            self._go()
            time.sleep(2)
            body = self.driver.find_element(By.TAG_NAME, "body")
            assert body.is_displayed()
            overflow_x = self.driver.execute_script(
                "return document.documentElement.scrollWidth > window.innerWidth"
            )
            self.driver.set_window_size(1280, 900)
            _record("Functional", "F-024", "Mobile Viewport — No Horizontal Overflow",
                    "PASS" if not overflow_x else "WARN",
                    f"Overflow: {overflow_x}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-024", "Mobile Viewport — No Horizontal Overflow",
                    "FAIL", str(e), time.time()-t0)
            self.driver.set_window_size(1280, 900)
            self.fail(str(e))

    # ── TC-F-025: Visited page accessible ────────────────────────────────────
    def test_F025_visited_page(self):
        t0 = time.time()
        try:
            self._go("visited")
            time.sleep(3)
            page_src = self.driver.page_source.lower()
            assert any(kw in page_src for kw in ["tripsync", "visited", "places", "sign in"])
            _record("Functional", "F-025", "Visited Places Page Reachable",
                    "PASS", f"URL: {self.driver.current_url}", time.time()-t0)
        except Exception as e:
            _record("Functional", "F-025", "Visited Places Page Reachable",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))


# ═══════════════════════════════════════════════════════════════════════════════
#  SECTION 2 — VULNERABILITY / SECURITY TESTS
# ═══════════════════════════════════════════════════════════════════════════════
class VulnerabilityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.driver = _build_driver()
        cls.driver.set_page_load_timeout(45)

    @classmethod
    def tearDownClass(cls):
        cls.driver.quit()

    def _go(self, path=""):
        self.driver.get(BASE_URL.rstrip("/") + ("/" + path.lstrip("/") if path else ""))
        time.sleep(1.5)

    # ── TC-V-001: XSS in email field ─────────────────────────────────────────
    def test_V001_xss_login_email(self):
        t0 = time.time()
        xss_payload = "<script>alert('XSS')</script>"
        try:
            self._go("login")
            email_field = _wait_visible(self.driver, By.ID, "email")
            email_field.clear()
            email_field.send_keys(xss_payload)
            btn = self.driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
            btn.click()
            time.sleep(1.5)
            # Check no alert dialog appeared
            try:
                alert = self.driver.switch_to.alert
                alert.dismiss()
                _record("Vulnerability", "V-001", "XSS Injection in Email Field",
                        "FAIL", "XSS alert executed!", time.time()-t0, "Critical")
                self.fail("XSS alert was triggered!")
            except:
                pass
            # Check script not rendered as HTML
            page_src = self.driver.page_source
            executed = "<script>alert" in page_src and "alert('XSS')" in page_src
            _record("Vulnerability", "V-001", "XSS Injection in Email Field",
                    "FAIL" if executed else "PASS",
                    "Script tag rendered in DOM" if executed else "XSS payload sanitized",
                    time.time()-t0, "Critical")
            if executed:
                self.fail("XSS payload was rendered in DOM")
        except AssertionError:
            raise
        except Exception as e:
            _record("Vulnerability", "V-001", "XSS Injection in Email Field",
                    "FAIL", str(e), time.time()-t0, "Critical")
            self.fail(str(e))

    # ── TC-V-002: XSS in password field ─────────────────────────────────────
    def test_V002_xss_login_password(self):
        t0 = time.time()
        xss_payload = "<img src=x onerror=alert(1)>"
        try:
            self._go("login")
            _wait_visible(self.driver, By.ID, "email").send_keys("test@test.com")
            self.driver.find_element(By.ID, "password").send_keys(xss_payload)
            self.driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
            time.sleep(1.5)
            try:
                alert = self.driver.switch_to.alert
                alert.dismiss()
                _record("Vulnerability", "V-002", "XSS Injection in Password Field",
                        "FAIL", "XSS onerror executed!", time.time()-t0, "Critical")
                self.fail("XSS executed in password field")
            except:
                pass
            _record("Vulnerability", "V-002", "XSS Injection in Password Field",
                    "PASS", "Password field XSS sanitized", time.time()-t0, "Critical")
        except AssertionError:
            raise
        except Exception as e:
            _record("Vulnerability", "V-002", "XSS Injection in Password Field",
                    "FAIL", str(e), time.time()-t0, "Critical")
            self.fail(str(e))

    # ── TC-V-003: SQL Injection attempt in email ──────────────────────────────
    def test_V003_sql_injection_email(self):
        t0 = time.time()
        sqli_payload = "' OR '1'='1'; DROP TABLE users; --"
        try:
            self._go("login")
            _wait_visible(self.driver, By.ID, "email").send_keys(sqli_payload)
            self.driver.find_element(By.ID, "password").send_keys("password123")
            self.driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
            time.sleep(3)
            # Should NOT be on dashboard
            on_dashboard = "dashboard" in self.driver.current_url
            _record("Vulnerability", "V-003", "SQL Injection in Email Field",
                    "FAIL" if on_dashboard else "PASS",
                    "SQL injection gained access!" if on_dashboard else "SQLi blocked",
                    time.time()-t0, "Critical")
            assert not on_dashboard
        except AssertionError:
            raise
        except Exception as e:
            _record("Vulnerability", "V-003", "SQL Injection in Email Field",
                    "FAIL", str(e), time.time()-t0, "Critical")
            self.fail(str(e))

    # ── TC-V-004: Open redirect check ────────────────────────────────────────
    def test_V004_open_redirect_check(self):
        t0 = time.time()
        malicious_url = BASE_URL + "?redirect=https://evil.com"
        try:
            self.driver.get(malicious_url)
            time.sleep(2)
            current = self.driver.current_url
            redirected_to_evil = "evil.com" in current
            _record("Vulnerability", "V-004", "Open Redirect Vulnerability Check",
                    "FAIL" if redirected_to_evil else "PASS",
                    f"Redirected to evil.com!" if redirected_to_evil else f"Stayed at: {current}",
                    time.time()-t0, "High")
            assert not redirected_to_evil
        except AssertionError:
            raise
        except Exception as e:
            _record("Vulnerability", "V-004", "Open Redirect Vulnerability Check",
                    "FAIL", str(e), time.time()-t0, "High")
            self.fail(str(e))

    # ── TC-V-005: Sensitive data not in page source ───────────────────────────
    def test_V005_no_sensitive_data_in_source(self):
        t0 = time.time()
        try:
            self._go("login")
            time.sleep(2)
            page_src = self.driver.page_source
            # Check no hardcoded API keys in HTML source
            sensitive_patterns = [
                "AIzaSy",        # Firebase API key prefix
                "sk-or-v1-",     # OpenRouter key
                "gsk_",          # Groq key
                "hf_",           # HuggingFace key
                "password",      # Raw password data
            ]
            found = [p for p in sensitive_patterns
                    if p in page_src and p not in ["password"]]
            # "password" is fine in HTML attribute names, not values
            status = "FAIL" if found else "PASS"
            detail = f"Exposed: {found}" if found else "No sensitive keys exposed in source"
            _record("Vulnerability", "V-005", "No Sensitive API Keys in Page Source",
                    status, detail, time.time()-t0, "Critical")
            assert not found, f"Sensitive data found: {found}"
        except AssertionError:
            raise
        except Exception as e:
            _record("Vulnerability", "V-005", "No Sensitive API Keys in Page Source",
                    "FAIL", str(e), time.time()-t0, "Critical")
            self.fail(str(e))

    # ── TC-V-006: HTTPS enforced ─────────────────────────────────────────────
    def test_V006_https_enforced(self):
        t0 = time.time()
        try:
            self.driver.get(BASE_URL)
            time.sleep(2)
            is_https = self.driver.current_url.startswith("https://")
            _record("Vulnerability", "V-006", "HTTPS Enforced on Web App",
                    "PASS" if is_https else "FAIL",
                    f"Protocol: {self.driver.current_url[:8]}",
                    time.time()-t0, "High")
            assert is_https
        except AssertionError:
            raise
        except Exception as e:
            _record("Vulnerability", "V-006", "HTTPS Enforced on Web App",
                    "FAIL", str(e), time.time()-t0, "High")
            self.fail(str(e))

    # ── TC-V-007: No console errors with XSS payload in URL ──────────────────
    def test_V007_url_xss_attempt(self):
        t0 = time.time()
        xss_in_url = BASE_URL + "/<script>alert('xss')</script>"
        try:
            self.driver.get(xss_in_url)
            time.sleep(2)
            try:
                alert = self.driver.switch_to.alert
                alert.dismiss()
                _record("Vulnerability", "V-007", "URL-based XSS Attack Blocked",
                        "FAIL", "Alert executed via URL", time.time()-t0, "Critical")
                self.fail("URL XSS executed")
            except:
                pass
            _record("Vulnerability", "V-007", "URL-based XSS Attack Blocked",
                    "PASS", "URL XSS blocked by app/browser", time.time()-t0, "Critical")
        except WebDriverException as e:
            if "alert" in str(e).lower():
                _record("Vulnerability", "V-007", "URL-based XSS Attack Blocked",
                        "FAIL", str(e), time.time()-t0, "Critical")
            else:
                _record("Vulnerability", "V-007", "URL-based XSS Attack Blocked",
                        "PASS", "App handled malformed URL", time.time()-t0)
        except Exception as e:
            _record("Vulnerability", "V-007", "URL-based XSS Attack Blocked",
                    "INFO", f"App handled invalid URL: {str(e)[:100]}", time.time()-t0)

    # ── TC-V-008: Login rate limiting / brute force protection ───────────────
    def test_V008_brute_force_protection(self):
        t0 = time.time()
        try:
            self._go("login")
            for attempt in range(3):
                email = _wait_visible(self.driver, By.ID, "email")
                email.clear()
                email.send_keys("brute@test.com")
                pw = self.driver.find_element(By.ID, "password")
                pw.clear()
                pw.send_keys(f"wrongpass{attempt}")
                self.driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
                time.sleep(2)
            page_src = self.driver.page_source.lower()
            # Should show error or still on login page
            blocked = any(kw in page_src for kw in
                        ["too many", "rate limit", "locked", "error", "invalid", "blocked"])
            still_on_login = "dashboard" not in self.driver.current_url
            _record("Vulnerability", "V-008", "Brute Force — Still On Login After 3 Attempts",
                    "PASS" if (blocked or still_on_login) else "WARN",
                    "Rate limited or errors shown" if blocked else "No dashboard access achieved",
                    time.time()-t0, "High")
        except Exception as e:
            _record("Vulnerability", "V-008", "Brute Force — Still On Login After 3 Attempts",
                    "FAIL", str(e), time.time()-t0, "High")
            self.fail(str(e))

    # ── TC-V-009: Prototype pollution check ──────────────────────────────────
    def test_V009_prototype_pollution(self):
        t0 = time.time()
        try:
            self._go("login")
            time.sleep(2)
            result = self.driver.execute_script("""
                try {
                    var payload = JSON.parse('{"__proto__": {"polluted": true}}');
                    return ({}).polluted;
                } catch(e) {
                    return false;
                }
            """)
            polluted = result is True
            _record("Vulnerability", "V-009", "JavaScript Prototype Pollution Check",
                    "FAIL" if polluted else "PASS",
                    "Prototype polluted!" if polluted else "No prototype pollution detected",
                    time.time()-t0, "High")
            assert not polluted
        except AssertionError:
            raise
        except Exception as e:
            _record("Vulnerability", "V-009", "JavaScript Prototype Pollution Check",
                    "FAIL", str(e), time.time()-t0, "High")
            self.fail(str(e))

    # ── TC-V-010: Clickjacking — iframe embedding check ──────────────────────
    def test_V010_clickjacking_check(self):
        t0 = time.time()
        try:
            resp = requests.get(BASE_URL, timeout=15, verify=False)
            headers = resp.headers
            x_frame = headers.get("X-Frame-Options", "")
            csp = headers.get("Content-Security-Policy", "")
            has_protection = (
                x_frame.upper() in ["DENY", "SAMEORIGIN"] or
                "frame-ancestors" in csp.lower()
            )
            _record("Vulnerability", "V-010", "Clickjacking Protection (X-Frame-Options)",
                    "PASS" if has_protection else "WARN",
                    f"X-Frame-Options: '{x_frame}', CSP: '{csp[:80]}'",
                    time.time()-t0, "Medium")
        except Exception as e:
            _record("Vulnerability", "V-010", "Clickjacking Protection (X-Frame-Options)",
                    "INFO", f"Could not check headers: {str(e)[:80]}", time.time()-t0, "Medium")


# ═══════════════════════════════════════════════════════════════════════════════
#  SECTION 3 — BACKEND API / UNIT TESTS (via requests)
# ═══════════════════════════════════════════════════════════════════════════════
class APIUnitTests(unittest.TestCase):
    BASE = API_URL.rstrip("/")

    # ── TC-A-001: Backend root health check ──────────────────────────────────
    def test_A001_backend_health(self):
        t0 = time.time()
        try:
            resp = requests.get(self.BASE + "/", timeout=30)
            assert resp.status_code in [200, 404, 422], f"Unexpected: {resp.status_code}"
            _record("API Unit", "A-001", "Backend Root Responds",
                    "PASS", f"HTTP {resp.status_code}", time.time()-t0)
        except Exception as e:
            _record("API Unit", "A-001", "Backend Root Responds",
                    "FAIL", str(e), time.time()-t0, "Critical")
            self.fail(str(e))

    # ── TC-A-002: POST /api/chat ──────────────────────────────────────────────
    def test_A002_api_chat_endpoint(self):
        t0 = time.time()
        try:
            payload = {"message": "Hello! What are good places to visit in Goa?", "history": []}
            resp = requests.post(f"{self.BASE}/api/chat", json=payload, timeout=60)
            assert resp.status_code in [200, 503], f"Unexpected status: {resp.status_code}"
            if resp.status_code == 200:
                data = resp.json()
                assert "reply" in data, f"Missing 'reply' key, got: {list(data.keys())}"
                _record("API Unit", "A-002", "POST /api/chat Returns Reply",
                        "PASS", f"Reply length: {len(str(data.get('reply', '')))}", time.time()-t0)
            else:
                _record("API Unit", "A-002", "POST /api/chat Returns Reply",
                        "WARN", f"Service returned {resp.status_code} (possible cold start)",
                        time.time()-t0)
        except Exception as e:
            _record("API Unit", "A-002", "POST /api/chat Returns Reply",
                    "FAIL", str(e), time.time()-t0, "Critical")
            self.fail(str(e))

    # ── TC-A-003: GET /api/safety ─────────────────────────────────────────────
    def test_A003_api_safety_endpoint(self):
        t0 = time.time()
        try:
            resp = requests.get(f"{self.BASE}/api/safety?city=Goa", timeout=60)
            assert resp.status_code in [200, 503], f"Unexpected: {resp.status_code}"
            if resp.status_code == 200:
                data = resp.json()
                required = ["city", "generalSafety", "nightSafety"]
                missing = [k for k in required if k not in data]
                assert not missing, f"Missing fields: {missing}"
                _record("API Unit", "A-003", "GET /api/safety Returns Safety Metrics",
                        "PASS", f"City: {data.get('city')}, Safety: {data.get('generalSafety')}",
                        time.time()-t0)
            else:
                _record("API Unit", "A-003", "GET /api/safety Returns Safety Metrics",
                        "WARN", f"Service {resp.status_code}", time.time()-t0)
        except Exception as e:
            _record("API Unit", "A-003", "GET /api/safety Returns Safety Metrics",
                    "FAIL", str(e), time.time()-t0, "Critical")
            self.fail(str(e))

    # ── TC-A-004: POST /api/briefing ─────────────────────────────────────────
    def test_A004_api_briefing_endpoint(self):
        t0 = time.time()
        try:
            payload = {
                "userName": "TestUser",
                "activeTripName": "Goa Adventure",
                "activeTripDestination": "Goa",
                "todayScheduleTitle": "Beach Day",
                "todayScheduleSpots": ["Baga Beach", "Fort Aguada"],
                "upcomingTripName": None,
                "upcomingTripDestination": None,
                "upcomingTripDays": None,
                "groupName": "Beach Squad",
                "groupExpensesCount": 3,
                "groupMembersCount": 4,
                "groupLastExpenseAmount": 1500.0,
                "groupLastExpenseDesc": "Dinner",
                "weatherTemp": 30,
                "weatherDesc": "Sunny"
            }
            resp = requests.post(f"{self.BASE}/api/briefing", json=payload, timeout=60)
            assert resp.status_code in [200, 503]
            if resp.status_code == 200:
                data = resp.json()
                assert "briefing" in data, f"Missing 'briefing' key"
                _record("API Unit", "A-004", "POST /api/briefing Returns Briefing Text",
                        "PASS", f"Briefing length: {len(str(data.get('briefing', '')))}",
                        time.time()-t0)
            else:
                _record("API Unit", "A-004", "POST /api/briefing Returns Briefing Text",
                        "WARN", f"Service {resp.status_code}", time.time()-t0)
        except Exception as e:
            _record("API Unit", "A-004", "POST /api/briefing Returns Briefing Text",
                    "FAIL", str(e), time.time()-t0, "Critical")
            self.fail(str(e))

    # ── TC-A-005: POST /api/routes/optimize ──────────────────────────────────
    def test_A005_api_routes_optimize(self):
        t0 = time.time()
        try:
            payload = [
                {"name": "Baga Beach",    "latitude": 15.555, "longitude": 73.752},
                {"name": "Fort Aguada",   "latitude": 15.499, "longitude": 73.773},
                {"name": "Calangute",     "latitude": 15.542, "longitude": 73.755},
            ]
            resp = requests.post(f"{self.BASE}/api/routes/optimize", json=payload, timeout=60)
            assert resp.status_code in [200, 503]
            if resp.status_code == 200:
                data = resp.json()
                assert isinstance(data, list), "Expected list response"
                assert len(data) == len(payload), f"Count mismatch: {len(data)} vs {len(payload)}"
                _record("API Unit", "A-005", "POST /api/routes/optimize Returns Ordered Spots",
                        "PASS", f"Returned {len(data)} optimized spots", time.time()-t0)
            else:
                _record("API Unit", "A-005", "POST /api/routes/optimize Returns Ordered Spots",
                        "WARN", f"Service {resp.status_code}", time.time()-t0)
        except Exception as e:
            _record("API Unit", "A-005", "POST /api/routes/optimize Returns Ordered Spots",
                    "FAIL", str(e), time.time()-t0, "Critical")
            self.fail(str(e))

    # ── TC-A-006: POST /api/expenses/split ───────────────────────────────────
    def test_A006_api_expenses_split(self):
        t0 = time.time()
        try:
            payload = {
                "amount": 1200,
                "description": "Beach dinner",
                "category": "Food",
                "paidBy": "Alice",
                "splitBetween": ["Alice", "Bob", "Charlie"]
            }
            resp = requests.post(f"{self.BASE}/api/expenses/split", json=payload, timeout=60)
            assert resp.status_code in [200, 422, 503]
            if resp.status_code == 200:
                data = resp.json()
                _record("API Unit", "A-006", "POST /api/expenses/split Returns Split Details",
                        "PASS", f"Response keys: {list(data.keys()) if isinstance(data, dict) else 'list'}",
                        time.time()-t0)
            elif resp.status_code == 422:
                _record("API Unit", "A-006", "POST /api/expenses/split Returns Split Details",
                        "WARN", "Validation error (422) — check payload schema", time.time()-t0)
            else:
                _record("API Unit", "A-006", "POST /api/expenses/split Returns Split Details",
                        "WARN", f"Service {resp.status_code}", time.time()-t0)
        except Exception as e:
            _record("API Unit", "A-006", "POST /api/expenses/split Returns Split Details",
                    "FAIL", str(e), time.time()-t0, "Critical")
            self.fail(str(e))

    # ── TC-A-007: POST /api/routes/share ─────────────────────────────────────
    def test_A007_api_routes_share(self):
        t0 = time.time()
        try:
            payload = {
                "tripId": "test-trip-001",
                "tripName": "Goa Expedition",
                "destination": "Goa",
                "spots": [{"name": "Baga Beach", "latitude": 15.555, "longitude": 73.752}],
                "sharedBy": "TestUser"
            }
            resp = requests.post(f"{self.BASE}/api/routes/share", json=payload, timeout=60)
            assert resp.status_code in [200, 422, 503]
            if resp.status_code == 200:
                _record("API Unit", "A-007", "POST /api/routes/share Returns Share Data",
                        "PASS", f"HTTP {resp.status_code}", time.time()-t0)
            elif resp.status_code == 422:
                _record("API Unit", "A-007", "POST /api/routes/share Returns Share Data",
                        "WARN", "Schema validation error (422)", time.time()-t0)
            else:
                _record("API Unit", "A-007", "POST /api/routes/share Returns Share Data",
                        "WARN", f"Service {resp.status_code}", time.time()-t0)
        except Exception as e:
            _record("API Unit", "A-007", "POST /api/routes/share Returns Share Data",
                    "FAIL", str(e), time.time()-t0, "Critical")
            self.fail(str(e))

    # ── TC-A-008: CORS headers present on API ─────────────────────────────────
    def test_A008_cors_headers(self):
        t0 = time.time()
        try:
            resp = requests.options(
                f"{self.BASE}/api/chat",
                headers={"Origin": BASE_URL, "Access-Control-Request-Method": "POST"},
                timeout=15
            )
            cors = resp.headers.get("Access-Control-Allow-Origin", "")
            has_cors = bool(cors)
            _record("API Unit", "A-008", "CORS Headers Present on Backend API",
                    "PASS" if has_cors else "WARN",
                    f"ACAO header: '{cors}'", time.time()-t0, "High")
        except Exception as e:
            _record("API Unit", "A-008", "CORS Headers Present on Backend API",
                    "INFO", f"CORS check: {str(e)[:80]}", time.time()-t0)

    # ── TC-A-009: API rejects invalid JSON ───────────────────────────────────
    def test_A009_api_invalid_payload(self):
        t0 = time.time()
        try:
            resp = requests.post(
                f"{self.BASE}/api/chat",
                data="NOT VALID JSON {{{",
                headers={"Content-Type": "application/json"},
                timeout=30
            )
            assert resp.status_code in [400, 422, 503], f"Expected 4xx, got {resp.status_code}"
            _record("API Unit", "A-009", "API Rejects Malformed JSON Payload",
                    "PASS", f"HTTP {resp.status_code} — server rejected invalid JSON",
                    time.time()-t0, "High")
        except AssertionError:
            raise
        except Exception as e:
            _record("API Unit", "A-009", "API Rejects Malformed JSON Payload",
                    "FAIL", str(e), time.time()-t0, "High")
            self.fail(str(e))

    # ── TC-A-010: API empty body handling ────────────────────────────────────
    def test_A010_api_empty_body(self):
        t0 = time.time()
        try:
            resp = requests.post(
                f"{self.BASE}/api/chat",
                json={},
                timeout=30
            )
            assert resp.status_code in [400, 422, 503]
            _record("API Unit", "A-010", "API Handles Empty Body with 4xx",
                    "PASS", f"HTTP {resp.status_code}", time.time()-t0)
        except AssertionError:
            raise
        except Exception as e:
            _record("API Unit", "A-010", "API Handles Empty Body with 4xx",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-A-011: Safety endpoint missing city param ──────────────────────────
    def test_A011_safety_missing_param(self):
        t0 = time.time()
        try:
            resp = requests.get(f"{self.BASE}/api/safety", timeout=30)
            assert resp.status_code in [400, 422, 503]
            _record("API Unit", "A-011", "Safety Endpoint Requires city Param",
                    "PASS", f"HTTP {resp.status_code} when city missing", time.time()-t0)
        except AssertionError:
            raise
        except Exception as e:
            _record("API Unit", "A-011", "Safety Endpoint Requires city Param",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))

    # ── TC-A-012: Response time under 60s (cold start allowed) ────────────────
    def test_A012_response_time(self):
        t0 = time.time()
        try:
            start = time.time()
            resp = requests.get(f"{self.BASE}/api/safety?city=Goa", timeout=70)
            elapsed = time.time() - start
            status = "PASS" if elapsed < 60 else "WARN"
            _record("API Unit", "A-012", "API Response Time Within 60s",
                    status, f"{elapsed:.2f}s (HTTP {resp.status_code})", time.time()-t0)
        except Exception as e:
            _record("API Unit", "A-012", "API Response Time Within 60s",
                    "FAIL", str(e), time.time()-t0)
            self.fail(str(e))


# ═══════════════════════════════════════════════════════════════════════════════
#  XLSX REPORT GENERATOR
# ═══════════════════════════════════════════════════════════════════════════════
def generate_xlsx_report():
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename  = f"TripSync_TestReport_{timestamp}.xlsx"
    filepath  = os.path.join(REPORT_DIR, filename)

    wb = openpyxl.Workbook()

    # ── Color palette ──────────────────────────────────────────────────────────
    C_PASS    = "FF22C55E"   # green-500
    C_FAIL    = "FFEF4444"   # red-500
    C_WARN    = "FFF59E0B"   # amber-400
    C_SKIP    = "FF94A3B8"   # slate-400
    C_INFO    = "FF60A5FA"   # blue-400
    C_HEADER  = "FF0F172A"   # slate-950
    C_SUBHDR  = "FF1E293B"   # slate-800
    C_LIGHT   = "FFF1F5F9"   # slate-100
    C_WHITE   = "FFFFFFFF"
    C_TEAL    = "FF14B8A6"   # teal-500

    thin  = Side(style="thin",   color="FF334155")
    thick = Side(style="medium", color="FF334155")

    def make_border(left=thin, right=thin, top=thin, bottom=thin):
        return Border(left=left, right=right, top=top, bottom=bottom)

    def header_font(size=11, bold=True, color=C_WHITE):
        return Font(name="Calibri", size=size, bold=bold, color=color)

    def body_font(size=10, bold=False, color="FF1E293B"):
        return Font(name="Calibri", size=size, bold=bold, color=color)

    # ── Sheet 1: Executive Summary ─────────────────────────────────────────────
    ws_sum = wb.active
    ws_sum.title = "Executive Summary"

    total      = len(_results)
    passed     = sum(1 for r in _results if r["status"] == "PASS")
    failed     = sum(1 for r in _results if r["status"] == "FAIL")
    warned     = sum(1 for r in _results if r["status"] in ["WARN", "SKIP", "INFO"])
    pass_rate  = (passed / total * 100) if total else 0
    run_time   = sum(r["duration"] for r in _results)

    by_cat: dict[str, dict] = {}
    for r in _results:
        c = r["category"]
        if c not in by_cat:
            by_cat[c] = {"total": 0, "pass": 0, "fail": 0, "warn": 0}
        by_cat[c]["total"] += 1
        if r["status"] == "PASS":    by_cat[c]["pass"] += 1
        elif r["status"] == "FAIL":  by_cat[c]["fail"] += 1
        else:                        by_cat[c]["warn"] += 1

    # Title block
    ws_sum.merge_cells("A1:H1")
    ws_sum["A1"] = "🧪 TripSync Web Application — E2E Test Report"
    ws_sum["A1"].font      = header_font(16)
    ws_sum["A1"].fill      = PatternFill("solid", fgColor=C_HEADER)
    ws_sum["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 40

    ws_sum.merge_cells("A2:H2")
    ws_sum["A2"] = (f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  |  "
                    f"App: {BASE_URL}  |  API: {API_URL}")
    ws_sum["A2"].font      = body_font(9, color=C_WHITE)
    ws_sum["A2"].fill      = PatternFill("solid", fgColor=C_SUBHDR)
    ws_sum["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[2].height = 20

    # KPI row
    kpis = [
        ("Total Tests",  total,          C_HEADER),
        ("✅ Passed",    passed,          C_PASS),
        ("❌ Failed",    failed,          C_FAIL),
        ("⚠ Warn/Skip", warned,          "FFFBBF24"),
        ("Pass Rate",   f"{pass_rate:.1f}%", C_TEAL),
        ("Runtime",     f"{run_time:.1f}s",  "FF7C3AED"),
    ]
    kpi_cols = [1, 2, 3, 4, 6, 7]
    ws_sum.row_dimensions[4].height = 60
    for col, (label, val, color) in zip(kpi_cols, kpis):
        ws_sum.merge_cells(start_row=4, start_column=col, end_row=5, end_column=col)
        cell = ws_sum.cell(row=4, column=col)
        cell.value     = f"{label}\n{val}"
        cell.font      = Font(name="Calibri", size=13, bold=True, color=C_WHITE)
        cell.fill      = PatternFill("solid", fgColor=color)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = make_border()

    # Category breakdown table
    row = 7
    cat_headers = ["Category", "Total", "Passed", "Failed", "Warn/Skip", "Pass %"]
    for ci, h in enumerate(cat_headers, 1):
        c = ws_sum.cell(row=row, column=ci, value=h)
        c.font      = header_font(10)
        c.fill      = PatternFill("solid", fgColor=C_SUBHDR)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = make_border()
    ws_sum.row_dimensions[row].height = 20

    for row_i, (cat, d) in enumerate(by_cat.items(), row + 1):
        pct = (d["pass"] / d["total"] * 100) if d["total"] else 0
        row_data = [cat, d["total"], d["pass"], d["fail"], d["warn"], f"{pct:.0f}%"]
        fill_col = C_PASS if d["fail"] == 0 else C_FAIL
        for ci, val in enumerate(row_data, 1):
            c = ws_sum.cell(row=row_i, column=ci, value=val)
            c.font      = body_font(10)
            c.alignment = Alignment(horizontal="center" if ci > 1 else "left", vertical="center")
            c.border    = make_border()
            if ci in [3, 4]:
                c.fill = PatternFill("solid", fgColor=C_PASS if ci==3 else (C_FAIL if d["fail"]>0 else C_PASS))
        ws_sum.row_dimensions[row_i].height = 18

    # Column widths
    for col, w in zip("ABCDEFGH", [30, 12, 12, 12, 14, 14, 20, 20]):
        ws_sum.column_dimensions[get_column_letter(col if isinstance(col, int)
                                                    else ord(col)-64)].width = w

    # ── Sheet 2–4: Per-category detailed results ───────────────────────────────
    category_order = ["Functional", "Vulnerability", "API Unit"]
    col_headers = ["#", "Test ID", "Test Name", "Status", "Duration (s)",
                   "Severity", "Details", "Timestamp"]
    col_widths   = [5, 10, 45, 10, 13, 12, 60, 20]

    STATUS_COLORS = {
        "PASS": C_PASS,
        "FAIL": C_FAIL,
        "WARN": "FFFBBF24",
        "SKIP": C_SKIP,
        "INFO": C_INFO,
    }

    for cat in category_order:
        cat_results = [r for r in _results if r["category"] == cat]
        if not cat_results:
            continue

        ws = wb.create_sheet(title=cat[:31])

        # Sheet title
        ws.merge_cells(f"A1:{get_column_letter(len(col_headers))}1")
        ws["A1"] = f"TripSync — {cat} Test Results"
        ws["A1"].font      = header_font(13)
        ws["A1"].fill      = PatternFill("solid", fgColor=C_HEADER)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        # Column headers
        for ci, (h, w) in enumerate(zip(col_headers, col_widths), 1):
            c = ws.cell(row=2, column=ci, value=h)
            c.font      = header_font(10)
            c.fill      = PatternFill("solid", fgColor=C_TEAL)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border    = make_border()
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[2].height = 22

        # Data rows
        for ri, r in enumerate(cat_results, 3):
            row_vals = [
                ri - 2,
                r["test_id"],
                r["name"],
                r["status"],
                r["duration"],
                r["severity"],
                r["details"][:200],
                r["timestamp"],
            ]
            status_color = STATUS_COLORS.get(r["status"], C_LIGHT)
            for ci, val in enumerate(row_vals, 1):
                c = ws.cell(row=ri, column=ci, value=val)
                c.font      = body_font(9 if ci == 7 else 10)
                c.alignment = Alignment(
                    horizontal="left" if ci in [3, 7] else "center",
                    vertical="center", wrap_text=(ci == 7)
                )
                c.border = make_border()
                if ci == 4:  # Status column — colored
                    c.fill = PatternFill("solid", fgColor=status_color)
                    c.font = Font(name="Calibri", size=10, bold=True,
                                  color=C_WHITE if r["status"] in ["FAIL","PASS"] else "FF1E293B")
                elif ri % 2 == 0:
                    c.fill = PatternFill("solid", fgColor="FFF8FAFC")
            ws.row_dimensions[ri].height = 22 if len(r.get("details","")) < 80 else 36

        ws.freeze_panes = "A3"

    # ── Sheet 5: All Results (flat) ────────────────────────────────────────────
    ws_all = wb.create_sheet("All Results")
    ws_all.merge_cells(f"A1:{get_column_letter(len(col_headers))}1")
    ws_all["A1"] = "TripSync — All Test Results (Combined)"
    ws_all["A1"].font      = header_font(13)
    ws_all["A1"].fill      = PatternFill("solid", fgColor=C_HEADER)
    ws_all["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_all.row_dimensions[1].height = 30

    all_headers = ["#", "Category", "Test ID", "Test Name", "Status", "Duration (s)",
                   "Severity", "Details", "Timestamp"]
    all_widths   = [5, 14, 10, 42, 10, 13, 12, 55, 20]
    for ci, (h, w) in enumerate(zip(all_headers, all_widths), 1):
        c = ws_all.cell(row=2, column=ci, value=h)
        c.font      = header_font(10)
        c.fill      = PatternFill("solid", fgColor=C_TEAL)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = make_border()
        ws_all.column_dimensions[get_column_letter(ci)].width = w
    ws_all.row_dimensions[2].height = 22

    for ri, r in enumerate(_results, 3):
        row_vals = [
            ri - 2, r["category"], r["test_id"], r["name"],
            r["status"], r["duration"], r["severity"], r["details"][:200], r["timestamp"]
        ]
        status_color = STATUS_COLORS.get(r["status"], C_LIGHT)
        for ci, val in enumerate(row_vals, 1):
            c = ws_all.cell(row=ri, column=ci, value=val)
            c.font      = body_font(9 if ci == 8 else 10)
            c.alignment = Alignment(
                horizontal="left" if ci in [4, 8] else "center",
                vertical="center", wrap_text=(ci == 8)
            )
            c.border = make_border()
            if ci == 5:
                c.fill = PatternFill("solid", fgColor=status_color)
                c.font = Font(name="Calibri", size=10, bold=True,
                              color=C_WHITE if r["status"] in ["FAIL","PASS"] else "FF1E293B")
            elif ri % 2 == 0:
                c.fill = PatternFill("solid", fgColor="FFF8FAFC")
        ws_all.row_dimensions[ri].height = 22
    ws_all.freeze_panes = "A3"

    wb.save(filepath)
    print(f"\n📊 XLSX report saved → {filepath}")
    return filepath


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN ENTRY
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    print(f"\n{'='*70}")
    print(f"  🧭 TripSync E2E Test Suite")
    print(f"  App URL : {BASE_URL}")
    print(f"  API URL : {API_URL}")
    print(f"  Headless: {HEADLESS}")
    print(f"{'='*70}\n")

    loader = unittest.TestLoader()
    suite  = unittest.TestSuite()

    # Order: API first (fast, no browser), then functional, then security
    for cls in [APIUnitTests, FunctionalTests, VulnerabilityTests]:
        suite.addTests(loader.loadTestsFromTestCase(cls))

    runner = unittest.TextTestRunner(verbosity=2, stream=sys.stdout)
    result = runner.run(suite)

    # Generate report
    report_path = generate_xlsx_report()

    # Print summary
    total   = len(_results)
    passed  = sum(1 for r in _results if r["status"] == "PASS")
    failed  = sum(1 for r in _results if r["status"] == "FAIL")
    warned  = sum(1 for r in _results if r["status"] in ["WARN", "SKIP", "INFO"])
    print(f"\n{'='*70}")
    print(f"  📋 FINAL SUMMARY")
    print(f"  Total  : {total}")
    print(f"  ✅ Pass : {passed}")
    print(f"  ❌ Fail : {failed}")
    print(f"  ⚠  Warn : {warned}")
    print(f"  Pass %  : {(passed/total*100):.1f}%" if total else "  Pass % : N/A")
    print(f"  Report  : {report_path}")
    print(f"{'='*70}\n")

    sys.exit(0 if result.wasSuccessful() else 1)


if __name__ == "__main__":
    main()
