import { describe, expect, it } from "vitest";
import { chatCompletionsUrl } from "../agent";

describe("chatCompletionsUrl", () => {
  it("appends chat completions to a /v1 base URL", () => {
    expect(chatCompletionsUrl("http://127.0.0.1:8646/v1")).toBe(
      "http://127.0.0.1:8646/v1/chat/completions"
    );
  });

  it("does not double-append chat completions", () => {
    expect(chatCompletionsUrl("http://127.0.0.1:8646/v1/chat/completions")).toBe(
      "http://127.0.0.1:8646/v1/chat/completions"
    );
  });

  it("routes legacy 8643 local URLs through the webpack dev proxy", () => {
    expect(chatCompletionsUrl("http://127.0.0.1:8643/v1")).toBe("/hermes/v1/chat/completions");
  });
});
