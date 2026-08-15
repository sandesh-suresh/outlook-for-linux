const UNREAD_OPTION_SELECTOR = '[role="option"][aria-label]';
const UNREAD_LABEL_PATTERN = /^unread\b/i;

/**
 * Derives the unread count from Outlook's message list.
 *
 * Each row in Outlook's mail list is a role="option" element whose
 * accessible name (aria-label) begins with the word "Unread" while the
 * message is unread. This counts matching rows without ever reading past
 * that prefix, so sender/subject/preview text is never inspected or
 * logged. Best-effort and DOM-shape-driven: Outlook's markup can change
 * without notice, so every failure path degrades to "no count" rather than
 * throwing.
 */
class UnreadListObserver {
  #lastCount = -1;

  init(config) {
    if (!config.useUnreadListLogic) {
      return;
    }
    console.debug("[UNREAD] Unread-count observer enabled");

    if (globalThis.document?.readyState === "loading") {
      globalThis.addEventListener("DOMContentLoaded", () =>
        this.#attachObserver(),
      );
    } else {
      this.#attachObserver();
    }
  }

  /**
   * @param {any} label - An option row's aria-label.
   * @returns {boolean} True when the label marks the row as unread.
   */
  isUnreadOptionLabel(label) {
    return typeof label === "string" && UNREAD_LABEL_PATTERN.test(label.trim());
  }

  /**
   * @param {string[]} labels - Aria-labels of the current option rows.
   * @returns {number} How many mark their row as unread.
   */
  countUnreadLabels(labels) {
    return labels.filter((label) => this.isUnreadOptionLabel(label)).length;
  }

  #countUnreadOptions() {
    const labels = [
      ...globalThis.document.querySelectorAll(UNREAD_OPTION_SELECTOR),
    ].map((element) => element.getAttribute("aria-label"));
    return this.countUnreadLabels(labels);
  }

  #handleMutations() {
    try {
      const count = this.#countUnreadOptions();
      if (count === this.#lastCount) {
        return;
      }
      this.#lastCount = count;

      console.debug(`[UNREAD] Unread count changed to ${count}`);
      globalThis.dispatchEvent(
        new CustomEvent("unread-count", { detail: { number: count } }),
      );
    } catch (error) {
      console.error("[UNREAD] Observer callback failed:", error.message);
    }
  }

  #attachObserver() {
    try {
      if (!globalThis.document?.body || !globalThis.MutationObserver) {
        console.error(
          "[UNREAD] DOM environment unsuitable, unread count disabled",
        );
        return;
      }

      const observer = new globalThis.MutationObserver(() =>
        this.#handleMutations(),
      );

      // childList/subtree catches rows appearing or disappearing (new mail,
      // virtualized scrolling); the aria-label attribute filter catches an
      // existing row flipping from unread to read in place.
      observer.observe(globalThis.document.body, {
        childList: true,
        subtree: true,
        attributes: true,
        attributeFilter: ["aria-label"],
      });

      // The list may already be fully rendered by the time this attaches, so
      // count once immediately rather than waiting for the next mutation.
      this.#handleMutations();
      console.debug("[UNREAD] Observer attached to document.body");
    } catch (error) {
      console.error("[UNREAD] Failed to attach observer:", error.message);
    }
  }
}

module.exports = new UnreadListObserver();
