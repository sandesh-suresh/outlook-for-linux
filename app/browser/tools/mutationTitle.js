const MAX_TITLE_LENGTH = 200;
const MAX_PLAUSIBLE_COUNT = 9999;
const UNREAD_MARKER = /^\((\d+)\)/;

/**
 * Derives the unread count from Outlook's page title.
 *
 * Outlook web prefixes the title with "(N)" when there are unread items, the
 * same convention Teams uses. This is best-effort: Outlook's DOM and title
 * format can change without notice, so every failure path degrades to
 * "no count" rather than throwing.
 */
class MutationObserverTitle {
  #lastNumber = -1;

  init(config) {
    if (!config.useMutationTitleLogic) {
      return;
    }
    console.debug("[TITLE] Unread-count observer enabled");

    if (globalThis.document?.readyState === "loading") {
      globalThis.addEventListener("DOMContentLoaded", () =>
        this.#attachObserver(),
      );
    } else {
      this.#attachObserver();
    }
  }

  /**
   * @param {any} title - The raw document title.
   * @returns {number|null} The unread count, or null when unusable.
   */
  parseUnreadCount(title) {
    if (typeof title !== "string") {
      return null;
    }

    const match = UNREAD_MARKER.exec(title.substring(0, MAX_TITLE_LENGTH));
    const number = match ? Number.parseInt(match[1], 10) : 0;

    if (Number.isNaN(number) || number < 0 || number > MAX_PLAUSIBLE_COUNT) {
      return null;
    }
    return number;
  }

  /** True when a mutation batch actually involves the <title> element. */
  #involvesTitleElement(mutations) {
    return mutations.some((mutation) => {
      const target = mutation.target;
      // characterData mutations target the text node inside <title>;
      // childList mutations on <title> target the element itself.
      if (
        target.nodeName === "TITLE" ||
        target.parentNode?.nodeName === "TITLE"
      ) {
        return true;
      }
      // <title> replaced outright, e.g. a React remount.
      const changedNodes = [...mutation.addedNodes, ...mutation.removedNodes];
      return changedNodes.some((node) => node.nodeName === "TITLE");
    });
  }

  #handleMutations(mutations) {
    try {
      if (!this.#involvesTitleElement(mutations)) {
        return;
      }

      // Never log the title itself: it carries the user's name and mailbox.
      const number = this.parseUnreadCount(globalThis.document.title);
      if (number === null || number === this.#lastNumber) {
        return;
      }
      this.#lastNumber = number;

      console.debug(`[TITLE] Unread count changed to ${number}`);
      globalThis.dispatchEvent(
        new CustomEvent("unread-count", { detail: { number } }),
      );
    } catch (error) {
      console.error("[TITLE] Observer callback failed:", error.message);
    }
  }

  #attachObserver() {
    try {
      if (!globalThis.document?.head || !globalThis.MutationObserver) {
        console.error(
          "[TITLE] DOM environment unsuitable, unread count disabled",
        );
        return;
      }

      const observer = new globalThis.MutationObserver((mutations) =>
        this.#handleMutations(mutations),
      );

      // Observing document.head with subtree survives Outlook replacing the
      // <title> element outright, and catches both childList (text node swap)
      // and characterData (in-place text edit) updates. Watching only the
      // current <title> with childList would miss in-place edits and leave the
      // badge stuck at a stale count.
      observer.observe(globalThis.document.head, {
        childList: true,
        characterData: true,
        subtree: true,
      });
      console.debug("[TITLE] Observer attached to document.head");
    } catch (error) {
      console.error("[TITLE] Failed to attach observer:", error.message);
    }
  }
}

module.exports = new MutationObserverTitle();
