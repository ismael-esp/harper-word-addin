# Third-party code

`word-addin/harper.bundle.js` and `harper-app/harper.bundle.js` are an unmodified build of these packages, bundled into one file with esbuild (see `harper-app/entry.js`):

| Package | Version | License | Copyright | License text |
|---|---|---|---|---|
| [harper.js](https://github.com/automattic/harper) (includes the Harper WebAssembly binary) | 3.0.0 | Apache-2.0 | Automattic / Elijah Potter | [harper-LICENSE.txt](harper-LICENSE.txt) |
| [fflate](https://github.com/101arrowz/fflate) | 0.8.3 | MIT | Arjun Barrett | [fflate-LICENSE.txt](fflate-LICENSE.txt) |

`dev/harper.d.ts` is harper.js's TypeScript type definitions (Apache-2.0), kept for reference.

Not included, loaded at run time: Microsoft's `office.js` (from Microsoft's CDN) and `office-addin-dev-certs` (downloaded by `Setup.cmd` through npm).
