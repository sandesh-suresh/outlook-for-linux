const APP_CLASS = "outlook-for-linux";

/**
 * Applies Chromium/Electron command-line switches derived from config.
 *
 * Must run before the first window is created.
 */
class CommandLineManager {
  /**
   * @param {object} config - The resolved startup config.
   * @param {object} electronApp - Electron's `app` object. Injectable for tests.
   */
  static applySwitches(config, electronApp) {
    // Both switches are needed for WM_CLASS to reach X11 and Wayland
    // compositors, which use it to match the window to its .desktop entry
    // (correct icon and taskbar grouping).
    electronApp.commandLine.appendSwitch("class", APP_CLASS);
    electronApp.commandLine.appendSwitch("wm-class", APP_CLASS);

    if (config.disableGpu) {
      console.info("[STARTUP] Disabling GPU acceleration");
      electronApp.commandLine.appendSwitch("disable-gpu");
      electronApp.commandLine.appendSwitch("disable-gpu-compositing");
      electronApp.commandLine.appendSwitch("disable-software-rasterizer");
    }
  }
}

module.exports = CommandLineManager;
