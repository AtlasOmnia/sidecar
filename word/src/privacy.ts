/**
 * Dynamic privacy classification based on the configured base URL.
 *
 * Derives privacy level at startup and whenever settings change.
 * Never hard-codes "Local" — always computed from the actual URL.
 *
 * Levels:
 * - local_bridge: 127.0.0.1, localhost, ::1
 * - private_network: RFC1918 IPv4 (10.x, 172.16-31.x, 192.168.x),
 *   link-local (169.254.x), IPv6 ULA (fc00::/7, fd00::/8),
 *   IPv6 link-local (fe80::/10)
 * - cloud: known cloud provider domains (suffix-matched, not substring)
 * - external: everything else
 */

export type PrivacyLevel = "local_bridge" | "private_network" | "cloud" | "external";

export function classifyPrivacy(baseUrl: string): PrivacyLevel {
  let hostPart: string;
  try {
    const parsed = new URL(baseUrl.trim());
    if (parsed.protocol !== "http:" && parsed.protocol !== "https:") return "external";
    hostPart = parsed.hostname.toLowerCase().replace(/\.$/, "");
  } catch {
    return "external";
  }

  // Loopback addresses
  if (
    hostPart === "127.0.0.1" ||
    hostPart === "localhost" ||
    hostPart === "[::1]" ||
    hostPart === "::1"
  ) {
    return "local_bridge";
  }

  // Private network: RFC1918 IPv4 and IPv4 link-local. Parse the entire
  // host as an address so lookalike DNS names such as 192.168.evil.example
  // cannot receive a misleading Private LAN label.
  const ipv4 = parseIpv4(hostPart);
  if (ipv4) {
    const [first, second] = ipv4;
    if (
      first === 10 ||
      (first === 172 && second >= 16 && second <= 31) ||
      (first === 192 && second === 168) ||
      (first === 169 && second === 254)
    ) {
      return "private_network";
    }
  }

  // IPv6 private ranges: ULA fc00::/7 and link-local fe80::/10
  if (hostPart.startsWith("[") && hostPart.endsWith("]")) {
    const ip = hostPart.slice(1, -1).toLowerCase();
    if (
      ip.startsWith("fc") || ip.startsWith("fd") ||
      ip.startsWith("fe8") || ip.startsWith("fe9") ||
      ip.startsWith("fea") || ip.startsWith("feb")
    ) {
      return "private_network";
    }
  }

  // Known cloud providers — SUFFIX-matched, not substring, to avoid spoofing
  // e.g. api.openai.com.attacker.example must NOT classify as OpenAI
  if (isKnownCloudHost(hostPart)) return "cloud";

  return "external";
}

/** Parse a canonical dotted-quad IPv4 host, rejecting DNS lookalikes. */
function parseIpv4(host: string): [number, number, number, number] | null {
  const parts = host.split(".");
  if (parts.length !== 4 || parts.some((part) => !/^\d{1,3}$/.test(part))) {
    return null;
  }
  const octets = parts.map(Number);
  if (octets.some((octet) => octet < 0 || octet > 255)) {
    return null;
  }
  return octets as [number, number, number, number];
}

/** Suffix-match against known cloud provider hostnames. */
function isKnownCloudHost(host: string): boolean {
  // Exact match or `.suffix` match only
  const known = [
    "api.openai.com",
    "openai.com",
    "api.anthropic.com",
    "anthropic.com",
    "api.deepseek.com",
    "deepseek.com",
    "openrouter.ai",
    "openrouter.io",
    "api.together.xyz",
    "api.groq.com",
    "groq.com",
    "api.mistral.ai",
    "mistral.ai",
  ];
  for (const k of known) {
    if (host === k || host.endsWith("." + k)) return true;
  }
  return false;
}

export const PRIVACY_LABELS: Record<
  PrivacyLevel,
  { text: string; cssClass: string; tooltip: string }
> = {
  local_bridge: {
    text: "Local bridge",
    cssClass: "privacy-local",
    tooltip:
      "Routing through local proxy (127.0.0.1). Your model backend may still send data to a cloud provider if configured to do so.",
  },
  private_network: {
    text: "Private LAN",
    cssClass: "privacy-lan",
    tooltip:
      "Routing to a private network address (RFC1918 / link-local / ULA). The immediate endpoint is on your local network, but a Hermes gateway at that address may still send data to a cloud provider if configured to do so.",
  },
  cloud: {
    text: "Cloud",
    cssClass: "privacy-cloud",
    tooltip:
      "Connected to a cloud AI provider. Document data is transmitted to external servers.",
  },
  external: {
    text: "External",
    cssClass: "privacy-external",
    tooltip:
      "Connected to a non-loopback URL. Data may leave your machine unless this is a local network address.",
  },
};
