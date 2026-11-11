/**
 * Dynamic privacy classification based on the configured base URL.
 *
 * Derives privacy level at startup and whenever settings change.
 * Never hard-codes "Local" — always computed from the actual URL.
 */

export type PrivacyLevel = "local_bridge" | "cloud" | "external";

export function classifyPrivacy(baseUrl: string): PrivacyLevel {
  const trimmed = baseUrl.trim().toLowerCase();
  // Remove protocol prefix for comparison
  const noProto = trimmed.replace(/^https?:\/\//, "");
  // Extract host:port — handle IPv6 brackets
  const hostPort = noProto.split("/")[0];
  let hostPart: string;
  if (hostPort.startsWith("[")) {
    // IPv6 address — everything up to the closing bracket
    const close = hostPort.indexOf("]");
    hostPart = close > 0 ? hostPort.slice(0, close + 1) : hostPort;
  } else {
    hostPart = hostPort.split(":")[0];
  }

  // Loopback addresses
  if (
    hostPart === "127.0.0.1" ||
    hostPart === "localhost" ||
    hostPart === "[::1]" ||
    hostPart === "0.0.0.0"
  ) {
    return "local_bridge";
  }

  // Known cloud providers
  if (
    hostPart.includes("api.openai.com") ||
    hostPart.includes("openai.com") ||
    hostPart.includes("anthropic.com") ||
    hostPart.includes("api.deepseek.com") ||
    hostPart.includes("deepseek.com") ||
    hostPart.includes("openrouter.ai") ||
    hostPart.includes("openrouter.io") ||
    hostPart.includes("api.together.xyz") ||
    hostPart.includes("groq.com") ||
    hostPart.includes("api.mistral.ai")
  ) {
    return "cloud";
  }

  return "external";
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
  cloud: {
    text: "Cloud",
    cssClass: "privacy-cloud",
    tooltip:
      "Connected to a cloud AI provider. Workbook data is transmitted to external servers.",
  },
  external: {
    text: "External",
    cssClass: "privacy-external",
    tooltip:
      "Connected to a non-loopback URL. Data may leave your machine unless this is a local network address.",
  },
};
