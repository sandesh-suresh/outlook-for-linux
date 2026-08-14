const TrayIconChooser = require("./trayIconChooser");

const CANVAS_SIZE = 140;
const BADGE_RADIUS = 40;
const BADGE_CENTRE_X = 100;
const BADGE_CENTRE_Y = 90;
const BADGE_TEXT_BASELINE_Y = 110;
const IMAGE_PNG = "image/png";

/**
 * Composites the unread count onto the tray icon and pushes it to the main
 * process. Runs in the renderer because it needs a canvas to draw the badge.
 */
class TrayIconRenderer {
  #lastRequestedCount;
  #updateSequence = 0;

  init(config, ipcRenderer) {
    this.ipcRenderer = ipcRenderer;
    this.config = config;

    // Required here rather than at module scope so the module can be loaded
    // outside Electron. Electron's entry point resolves to the installed
    // binary and downloads it when missing, which would make the unit tests
    // fetch ~100 MB under CI's `npm ci --ignore-scripts`.
    const { nativeImage } = require("electron");

    const iconChooser = new TrayIconChooser(config);
    this.baseIcon = nativeImage.createFromPath(iconChooser.getFile());
    this.iconSize = this.baseIcon.getSize();

    globalThis.addEventListener(
      "unread-count",
      this.updateActivityCount.bind(this),
    );
    console.debug("[TRAY] Badge renderer initialised");
  }

  /** Clears deduplication state. Test seam only. */
  resetForTests() {
    this.#lastRequestedCount = undefined;
    this.#updateSequence = 0;
  }

  async updateActivityCount(event) {
    const count = event.detail.number;

    // Deduplicate against the most recently *requested* count, not the last
    // one that finished sending: comparing against the completed count would
    // swallow a clear-to-zero that arrives while a non-zero render is still
    // in flight, leaving the badge stuck.
    if (count === this.#lastRequestedCount) {
      return;
    }
    this.#lastRequestedCount = count;

    // Each update takes a sequence token, so an update whose render finishes
    // after a newer one started is discarded rather than overwriting it.
    const sequence = ++this.#updateSequence;

    // A zero count needs no canvas work: main falls back to the base icon.
    let icon = null;
    if (count > 0) {
      try {
        icon = await this.render(count);
      } catch (error) {
        console.error("[TRAY] Icon render failed:", error.message);
        // Allow a later event with the same count to retry.
        this.#lastRequestedCount = undefined;
        return;
      }
    }

    if (sequence !== this.#updateSequence) {
      return;
    }

    this.ipcRenderer.send("tray-update", {
      icon,
      flash: count > 0 && !this.config.disableNotificationWindowFlash,
      count,
    });

    if (!this.config.disableBadgeCount) {
      await this.ipcRenderer
        .invoke("set-badge-count", count)
        .catch((err) =>
          console.error("[TRAY] Failed to set badge count:", err.message),
        );
    }
  }

  /**
   * @param {number} count - The unread count to draw.
   * @returns {Promise<string>} A PNG data URL. Never rejects: on any drawing
   *   failure it resolves with the unbadged base icon.
   */
  render(count) {
    return new Promise((resolve) => {
      const baseIconData = this.baseIcon.toDataURL(IMAGE_PNG);
      if (!baseIconData || baseIconData === "data:,") {
        console.error("[TRAY] Base icon produced no data");
        resolve(baseIconData);
        return;
      }

      const canvas = document.createElement("canvas");
      canvas.width = CANVAS_SIZE;
      canvas.height = CANVAS_SIZE;

      const image = new Image();
      image.onerror = () => {
        console.error("[TRAY] Base icon failed to load for rendering");
        resolve(baseIconData);
      };
      image.onload = () => {
        try {
          resolve(this.#drawBadge(canvas, image, count));
        } catch (error) {
          console.error("[TRAY] Canvas drawing failed:", error.message);
          resolve(baseIconData);
        }
      };
      image.src = baseIconData;
    });
  }

  #drawBadge(canvas, image, count) {
    const ctx = canvas.getContext("2d");
    ctx.drawImage(image, 0, 0, CANVAS_SIZE, CANVAS_SIZE);

    if (count > 0 && !this.config.disableBadgeCount) {
      ctx.fillStyle = "red";
      ctx.beginPath();
      ctx.ellipse(
        BADGE_CENTRE_X,
        BADGE_CENTRE_Y,
        BADGE_RADIUS,
        BADGE_RADIUS,
        0,
        0,
        2 * Math.PI,
      );
      ctx.fill();

      ctx.textAlign = "center";
      ctx.fillStyle = "white";
      ctx.font =
        'bold 70px "Segoe UI","Helvetica Neue",Helvetica,Arial,sans-serif';
      ctx.fillText(
        count > 9 ? "+" : count.toString(),
        BADGE_CENTRE_X,
        BADGE_TEXT_BASELINE_Y,
      );
    }

    return this.#resizeToIconSize(canvas).toDataURL();
  }

  #resizeToIconSize(canvas) {
    const resized = document.createElement("canvas");
    resized.width = this.iconSize.width;
    resized.height = this.iconSize.height;

    const ctx = resized.getContext("2d");
    ctx.scale(
      this.iconSize.width / canvas.width,
      this.iconSize.height / canvas.height,
    );
    ctx.drawImage(canvas, 0, 0);

    return resized;
  }
}

module.exports = new TrayIconRenderer();
