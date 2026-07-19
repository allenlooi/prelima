// Minimal GA4 wrapper. initGA() is only ever called from main.jsx (the real
// Vite entry), so this file has no build-time dependency (import.meta.env,
// bundler-only syntax) and stays safe to import from the no-build dev harness.
let enabled = false;

export function initGA(measurementId) {
  if (!measurementId || typeof window === "undefined" || enabled) return;
  enabled = true;

  const script = document.createElement("script");
  script.async = true;
  script.src = `https://www.googletagmanager.com/gtag/js?id=${measurementId}`;
  document.head.appendChild(script);

  window.dataLayer = window.dataLayer || [];
  function gtag() { window.dataLayer.push(arguments); }
  window.gtag = gtag;
  gtag("js", new Date());
  // The app is a single-page flow with client-side view changes, not full
  // navigations — we fire page_view ourselves on each view change instead
  // of letting gtag auto-fire once on initial script load.
  gtag("config", measurementId, { send_page_view: false });
}

export function pageview(path) {
  if (typeof window === "undefined" || typeof window.gtag !== "function") return;
  window.gtag("event", "page_view", {
    page_path: path,
    page_location: window.location.href,
    page_title: document.title,
  });
}

export function trackEvent(name, params = {}) {
  if (typeof window === "undefined" || typeof window.gtag !== "function") return;
  window.gtag("event", name, params);
}
