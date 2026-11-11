export type PrivacyClass = "local" | "cloud" | "external";
export const PRIVACY_LABELS: Record<PrivacyClass, { label: string; tooltip: string }> = {
  local: { label: "Local bridge", tooltip: "Requests use the loopback bridge; the configured backend may still be cloud-hosted." },
  cloud: { label: "Cloud", tooltip: "Requests are sent to a known cloud model provider." },
  external: { label: "External", tooltip: "Requests are sent to a non-loopback external endpoint." },
};
export function classifyPrivacy(baseUrl: string): PrivacyClass {
  try {
    const url = new URL(baseUrl);
    if (url.hostname === "127.0.0.1" || url.hostname === "localhost" || url.hostname === "::1" || url.hostname === "[::1]") return "local";
    if (["api.openai.com", "api.anthropic.com", "api.x.ai", "generativelanguage.googleapis.com"].includes(url.hostname.toLowerCase())) return "cloud";
  } catch { return "external"; }
  return "external";
}
