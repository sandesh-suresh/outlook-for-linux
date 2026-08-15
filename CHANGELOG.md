# Changelog

## [0.3.2](https://github.com/sandesh-suresh/outlook-for-linux/compare/v0.3.1...v0.3.2) (2026-08-15)


### Bug Fixes

* **deps:** pin electron to 43.2.0, 43.3+ breaks GNOME tray icon (elec… ([d523191](https://github.com/sandesh-suresh/outlook-for-linux/commit/d52319165f9f95c431db67b1bee324401a01c351))
* **deps:** pin electron to 43.2.0, 43.3+ breaks GNOME tray icon (electron[#52674](https://github.com/sandesh-suresh/outlook-for-linux/issues/52674)) ([50751ee](https://github.com/sandesh-suresh/outlook-for-linux/commit/50751ee2b0239b6a327d88a5792987b29a8903d3))

## [0.3.1](https://github.com/sandesh-suresh/outlook-for-linux/compare/v0.3.0...v0.3.1) (2026-08-15)


### Maintenance

* **deps-dev:** bump electron from 42.8.1 to 43.4.0 ([b2f9ca0](https://github.com/sandesh-suresh/outlook-for-linux/commit/b2f9ca0670c4aceaf3da3328a7cff64151763c99))
* **deps-dev:** bump electron from 42.8.1 to 43.4.0 ([a82e956](https://github.com/sandesh-suresh/outlook-for-linux/commit/a82e956ec978cfb1854795d4db46f26a72de13f5))

## [0.3.0](https://github.com/sandesh-suresh/outlook-for-linux/compare/v0.2.3...v0.3.0) (2026-08-15)


### Features

* **tray:** add unread-count menu item and fix unread detection ([775d24d](https://github.com/sandesh-suresh/outlook-for-linux/commit/775d24d48972fca49aba085fac62931a1a9f94f3))
* **tray:** add unread-count menu item and fix unread detection ([d59fe19](https://github.com/sandesh-suresh/outlook-for-linux/commit/d59fe19d159072c8497983c1b65a303beb01f7ce))

## [0.2.3](https://github.com/sandesh-suresh/outlook-for-linux/compare/v0.2.2...v0.2.3) (2026-08-15)


### Bug Fixes

* **notifications:** forward Outlook's in-app toast to native notifica… ([700d5d9](https://github.com/sandesh-suresh/outlook-for-linux/commit/700d5d943f7572464ec180c06a59f2f4b5c4ebd6))
* **notifications:** forward Outlook's in-app toast to native notifications ([5d91ee2](https://github.com/sandesh-suresh/outlook-for-linux/commit/5d91ee2bbf19049c701a4266bccae452572f3cde))

## [0.2.2](https://github.com/sandesh-suresh/outlook-for-linux/compare/v0.2.1...v0.2.2) (2026-08-14)


### Bug Fixes

* **packaging:** keep /usr/lib/.build-id links out of the rpm ([7852e8e](https://github.com/sandesh-suresh/outlook-for-linux/commit/7852e8e5e7e40a811f1aada8f4dc811702f0f088))
* **packaging:** keep /usr/lib/.build-id links out of the rpm ([81c1b76](https://github.com/sandesh-suresh/outlook-for-linux/commit/81c1b762be1bab1acdb43d36a2ab5d4424871735))

## [0.2.1](https://github.com/sandesh-suresh/outlook-for-linux/compare/v0.2.0...v0.2.1) (2026-08-14)


### Bug Fixes

* **ci:** attach release packages from the release workflow ([e3f2349](https://github.com/sandesh-suresh/outlook-for-linux/commit/e3f234948276aa945f942ec46d6be2f08f7e25f3))


### CI/CD

* leave code scanning to GitHub's default CodeQL setup ([6ef36ba](https://github.com/sandesh-suresh/outlook-for-linux/commit/6ef36bab0c8b01554d42d32debb587075627dc25))

## [0.2.0](https://github.com/sandesh-suresh/outlook-for-linux/compare/v0.1.0...v0.2.0) (2026-08-14)


### Features

* **config:** load configuration with defaults, file and flag precedence ([075d0a6](https://github.com/sandesh-suresh/outlook-for-linux/commit/075d0a6caa5fc6b294c149e309eebff97cd05df6))
* **notifications:** render Outlook's notifications natively ([532669d](https://github.com/sandesh-suresh/outlook-for-linux/commit/532669d986c3490399a1b30cc1878094f383c8b7))
* **security:** add the IPC allowlist, network classifier and CLI switches ([6fba79b](https://github.com/sandesh-suresh/outlook-for-linux/commit/6fba79befad6b3c4a9bf2e05abd155c7dc9fafc1))
* **tray:** show unread counts in the tray ([89c6866](https://github.com/sandesh-suresh/outlook-for-linux/commit/89c68662f09f0b3812b501c5405df27c93ca4b32))
* wire up the window, IPC registration and main process entry ([caa40ea](https://github.com/sandesh-suresh/outlook-for-linux/commit/caa40ea161c3194d2d9d2183cd1d835ba2721781))


### Bug Fixes

* **tray:** require electron lazily in the badge renderer ([0796622](https://github.com/sandesh-suresh/outlook-for-linux/commit/07966225542122f9571f675bb22939bb18421a49))
* **tray:** require electron lazily in the badge renderer ([26e9313](https://github.com/sandesh-suresh/outlook-for-linux/commit/26e93132f9303f099f7d39408be238cace13fcc3))


### Documentation

* add MVP design spec for outlook-for-linux ([64bbc56](https://github.com/sandesh-suresh/outlook-for-linux/commit/64bbc56e9b13e552a62093ceb6ed3b838ed91020))
* document configuration and packaging ([7c2b153](https://github.com/sandesh-suresh/outlook-for-linux/commit/7c2b153ec94ccad2de65f82913e215d57a9cc613))


### Maintenance

* scaffold the outlook-for-linux Electron project ([785e3d4](https://github.com/sandesh-suresh/outlook-for-linux/commit/785e3d431d48638e3e779a04494ae0d7773bbb4b))
