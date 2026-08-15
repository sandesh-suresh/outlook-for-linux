/**
 * Forwards Outlook's in-app ("Alert" style) new-mail toast to the native
 * notification pipeline.
 *
 * Outlook's "Alert" notification style never touches window.Notification or
 * the Push API — it renders its own DOM toast and announces it through an
 * aria-live="polite" region for screen readers. This watches that region and
 * replays the announcement through the same Notification shim used for the
 * standards-based path, so both styles end up delivered as native OS
 * notifications. Best-effort and DOM-shape-driven: Outlook's markup can
 * change without notice, so every failure path degrades to "no notification"
 * rather than throwing.
 */
const seenText = new WeakMap();

// Icon-font glyphs and unread badges show up as their own tiny text runs
// (observed: 2 characters) alongside the real sender/subject/preview leaves.
// Real sender names and content are never this short, so runs below this
// length are treated as decorative noise, not content.
const MIN_MEANINGFUL_RUN_LENGTH = 4;

/**
 * Builds a title/body pair from the live region's text-bearing leaves, in
 * document order. The region has no separators between elements, so once
 * decorative icon/badge runs are filtered out, the first remaining leaf is
 * the sender name and the rest is the subject/preview, which this joins
 * onto separate lines in the body.
 *
 * @param {string[]|undefined} runs - Trimmed leaf text runs, in order.
 * @param {string} appTitle - Fallback title when there is no separate sender.
 * @returns {{title: string, body: string}|null} The message, or null.
 */
function buildToastMessage(runs, appTitle) {
  if (!Array.isArray(runs)) {
    return null;
  }

  const cleaned = runs
    .map((run) => run.trim())
    .filter((run) => run.length >= MIN_MEANINGFUL_RUN_LENGTH);
  if (cleaned.length === 0) {
    return null;
  }

  if (cleaned.length === 1) {
    return { title: appTitle, body: cleaned[0] };
  }

  const [title, ...rest] = cleaned;
  return { title, body: rest.join("\n") };
}

class ToastObserver {
  init(config) {
    console.debug("[TOAST] Observer enabled");

    if (globalThis.document?.readyState === "loading") {
      globalThis.addEventListener("DOMContentLoaded", () =>
        this.#attachObserver(config),
      );
    } else {
      this.#attachObserver(config);
    }
  }

  // The whole toast is itself one clickable BUTTON, so a plain "skip every
  // button" rule would exclude all of its content. Only a button NESTED
  // inside that outer button — an action control like dismiss, carrying its
  // own accessible-name text (e.g. Outlook's "Microsoft Outlook" label) — is
  // skipped; the outer wrapper is still walked normally.
  #collectTextRuns(node, out, insideButton = false) {
    if (!(node instanceof globalThis.Element)) {
      return;
    }
    const isButton = node.tagName === "BUTTON";
    if (isButton && insideButton) {
      return;
    }
    let ownText = "";
    for (const child of node.childNodes) {
      if (child.nodeType === 3) {
        ownText += child.textContent;
      }
    }
    if (ownText.trim() !== "") {
      out.push(ownText);
    }
    for (const child of node.children) {
      this.#collectTextRuns(child, out, insideButton || isButton);
    }
  }

  #handleLiveRegionText(node, config) {
    try {
      const text = node.textContent || "";
      const last = seenText.get(node);
      seenText.set(node, text);

      if (text.trim() === "" || text === last) {
        return;
      }

      const runs = [];
      this.#collectTextRuns(node, runs);
      const message = buildToastMessage(runs, config.appTitle);
      if (!message) {
        return;
      }

      console.debug("[TOAST] Live region announced a new toast");
      new globalThis.Notification(message.title, { body: message.body });
    } catch (error) {
      console.error("[TOAST] Failed to handle live region text:", error.message);
    }
  }

  #watch(node, config) {
    if (seenText.has(node) || !globalThis.MutationObserver) {
      return;
    }
    seenText.set(node, "");
    const regionObserver = new globalThis.MutationObserver(() =>
      this.#handleLiveRegionText(node, config),
    );
    regionObserver.observe(node, {
      childList: true,
      characterData: true,
      subtree: true,
    });
  }

  #isPoliteLiveRegion(node) {
    return (
      node instanceof globalThis.Element &&
      node.getAttribute("aria-live") === "polite"
    );
  }

  #attachObserver(config) {
    try {
      if (!globalThis.document?.body || !globalThis.MutationObserver) {
        console.error("[TOAST] DOM environment unsuitable, toast forwarding disabled");
        return;
      }

      const observer = new globalThis.MutationObserver((mutations) => {
        for (const mutation of mutations) {
          if (mutation.type !== "childList") {
            continue;
          }
          for (const node of mutation.addedNodes) {
            if (this.#isPoliteLiveRegion(node)) {
              this.#watch(node, config);
            }
          }
        }
      });

      observer.observe(globalThis.document.body, {
        childList: true,
        subtree: true,
      });

      for (const node of globalThis.document.body.querySelectorAll(
        '[aria-live="polite"]',
      )) {
        this.#watch(node, config);
      }

      console.debug("[TOAST] Observer attached to document.body");
    } catch (error) {
      console.error("[TOAST] Failed to attach observer:", error.message);
    }
  }
}

const instance = new ToastObserver();
instance.buildToastMessage = buildToastMessage;

module.exports = instance;
