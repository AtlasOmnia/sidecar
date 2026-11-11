import https from "node:https";

/**
 * Upper bound on response bytes before the request is destroyed.
 * Shared between the runner and any other consumer of requestText.
 */
export const DEFAULT_MAX_RESPONSE_BYTES = 2 * 1024 * 1024;

/**
 * Fetch a URL as text with an absolute start-to-finish deadline.
 *
 * A trickling or stalled server cannot defeat this: once `timeout` ms
 * have elapsed the underlying request is destroyed and the promise
 * resolves `null`.  Late events after settlement are ignored and
 * cleanup (clearTimeout) is guaranteed exactly once.
 *
 * @param {string|URL} url
 * @param {{ timeout?: number, maxBytes?: number, transport?: typeof https }} [options]
 * @returns {Promise<string|null>}
 */
export function requestText(url, { timeout = 2_000, maxBytes = DEFAULT_MAX_RESPONSE_BYTES, transport = https } = {}) {
  return new Promise((resolve) => {
    let settled = false;
    const done = (value) => {
      if (settled) return;
      settled = true;
      clearTimeout(deadline);
      resolve(value);
    };

    // Absolute deadline — a trickling server cannot defeat this.
    const deadline = setTimeout(() => {
      request.destroy();
      done(null);
    }, timeout);

    const request = transport.get(
      url,
      { rejectUnauthorized: false },
      (response) => {
        if (response.statusCode !== 200) {
          response.resume();
          done(null);
          return;
        }
        const chunks = [];
        let byteLength = 0;
        response.on("data", (chunk) => {
          byteLength += chunk.length;
          if (byteLength > maxBytes) {
            request.destroy();
            done(null);
            return;
          }
          chunks.push(chunk);
        });
        response.once("end", () => {
          done(Buffer.concat(chunks).toString("utf8"));
        });
      },
    );
    request.once("error", () => done(null));
  });
}

/**
 * Construct the environment object for the acceptance dev-server child process.
 *
 * The returned object is a shallow merge of `baseEnv` with the required
 * acceptance keys.  It deliberately does NOT inject a machine-specific
 * `PYTHON` key.
 *
 * @param {Record<string,string|undefined>} [baseEnv]
 * @returns {Record<string,string|undefined>}
 */
export function buildRunnerEnv(baseEnv = process.env) {
  return {
    ...baseEnv,
    NODE_ENV: "development",
    HERMES_OFFICE_ACCEPTANCE: "1",
  };
}
