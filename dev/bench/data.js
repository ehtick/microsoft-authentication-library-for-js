window.BENCHMARK_DATA = {
  "lastUpdate": 1756178084172,
  "repoUrl": "https://github.com/ehtick/microsoft-authentication-library-for-js",
  "entries": {
    "msal-node client-credential Regression Test": [
      {
        "commit": {
          "author": {
            "email": "dasau@microsoft.com",
            "name": "Dan Saunders",
            "username": "codexeon"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "2d931672bcdd0e18ac6b780971d8029ef7cb7872",
          "message": "Fix exception when using claims with Nested App Auth in JS Runtime environment (#7926)\n\nJS Runtime does not support full crypto API, and is running into an\nexception in hydrateCache while trying to generate a sha-256 hash of the\nclaims. The token response is received, but never returned to developer\nbecause writing to cache failed. For this scenario, there is already a\ncache in the host, so it is still able to avoid a network call if a\nprevious request with the same claim was made.",
          "timestamp": "2025-07-14T12:35:08-07:00",
          "tree_id": "545e3b391595e3034c2474db680a56747422da46",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/2d931672bcdd0e18ac6b780971d8029ef7cb7872"
        },
        "date": 1752551434521,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 251764,
            "range": "±0.87%",
            "unit": "ops/sec",
            "extra": "224 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 251496,
            "range": "±0.96%",
            "unit": "ops/sec",
            "extra": "222 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "shylasummers@users.noreply.github.com",
            "name": "shylasummers",
            "username": "shylasummers"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "bc5f2a5a1646de3a71853f7d09b44eb4f65306ad",
          "message": "Bump MSAL Browser to 4.17.0 (#7951)",
          "timestamp": "2025-07-29T16:48:15-07:00",
          "tree_id": "59cc06598ac346efd97aa9fad1c156ce28bfe8bb",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/bc5f2a5a1646de3a71853f7d09b44eb4f65306ad"
        },
        "date": 1753849496544,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 244656,
            "range": "±0.81%",
            "unit": "ops/sec",
            "extra": "235 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 243134,
            "range": "±0.70%",
            "unit": "ops/sec",
            "extra": "230 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "ydi.w127@gmail.com",
            "name": "Yongdi Wang",
            "username": "yongdiw"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "2ffb0a2d38a153a952e9a1522f4aa4d1e503de2a",
          "message": "Add support for custom claims and password change required error (#7948)\n\nThis pull request introduces support for custom claims in various\nauthentication flows within the `msal-browser` library. The most\nsignificant changes include adding a `claims` field to multiple input\ntypes, ensuring claims are valid JSON strings, and propagating the\n`claims` field through the authentication process.\n\n### Support for custom claims:\n\n* Added a `claims` field to several input types (`SignInInputs`,\n`ResetPasswordInputs`, `AccessTokenRetrievalInputs`, and others) to\nallow custom claims during authentication.\n\n### Validation enhancements:\n\n* Introduced a new utility function, `ensureArgumentIsJSONString`, to\nvalidate that the `claims` field is a properly formatted JSON string.\nThis function is used in multiple places to ensure input integrity.\n\n### Integration into authentication flows:\n\n* Updated the `CustomAuthStandardController` and related classes to\nhandle the `claims` field during sign-in and token retrieval processes.\n* Modified API request types and parameters to include the `claims`\nfield, ensuring it is passed to the backend during token requests.\n### Error handling improvements:\n\n* Added methods to detect specific errors, such as password reset\nrequirements, during authentication flows.\n\n### Unit testing:\n\n* Enhanced the `ArgumentValidator` unit tests to cover the new\n`ensureArgumentIsJSONString` function.",
          "timestamp": "2025-07-30T09:33:33+01:00",
          "tree_id": "9c1147ea6f99b6e162fbea244def2ea601bb324d",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/2ffb0a2d38a153a952e9a1522f4aa4d1e503de2a"
        },
        "date": 1753866884734,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 250904,
            "range": "±0.70%",
            "unit": "ops/sec",
            "extra": "234 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 248049,
            "range": "±0.72%",
            "unit": "ops/sec",
            "extra": "234 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "kshabelko@microsoft.com",
            "name": "Konstantin Shabelko",
            "username": "konstantin-msft"
          },
          "committer": {
            "email": "kshabelko@microsoft.com",
            "name": "Konstantin",
            "username": "konstantin-msft"
          },
          "distinct": true,
          "id": "ebef0077979a685e9b6655eca7bd63e578873367",
          "message": "- Fix CVEs",
          "timestamp": "2025-08-04T17:58:37-04:00",
          "tree_id": "52e36f5b1dfeb3e330f0faeeb7b13ca9efa18a06",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/ebef0077979a685e9b6655eca7bd63e578873367"
        },
        "date": 1754367605766,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 242909,
            "range": "±1.06%",
            "unit": "ops/sec",
            "extra": "210 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 241914,
            "range": "±0.87%",
            "unit": "ops/sec",
            "extra": "232 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "152663010+msal-js-release-automation[bot]@users.noreply.github.com",
            "name": "msal-js-release-automation[bot]",
            "username": "msal-js-release-automation[bot]"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "c7b8239b5921d11faf4ac230097f96fdc55ab585",
          "message": "Release PR: official (#7969)\n\nThis PR contains the changelogs and version bumps for the MSAL.js 3P\nreleases.\n\nCo-authored-by: MSAL.js Release Automation <msaljsbuilds@microsoft.com>",
          "timestamp": "2025-08-05T15:37:57-07:00",
          "tree_id": "3e36ab6190db2a4e59697db7ec8d218f75b62bf1",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/c7b8239b5921d11faf4ac230097f96fdc55ab585"
        },
        "date": 1754450534841,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 246755,
            "range": "±0.76%",
            "unit": "ops/sec",
            "extra": "234 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 249566,
            "range": "±0.67%",
            "unit": "ops/sec",
            "extra": "224 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "lalima.sharda@gmail.com",
            "name": "Lalima Sharda",
            "username": "lalimasharda"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "5dc2f555537542d054dcc944e9b3b6413ff65754",
          "message": "getNativeAccountId bug fix (#7960)\n\nFall back to getting the native account id of the active account if no\nloginhint or sid is provided.",
          "timestamp": "2025-08-06T09:20:07-07:00",
          "tree_id": "3e459a084cf1aae337d01ebaf3f05aa1ce633397",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/5dc2f555537542d054dcc944e9b3b6413ff65754"
        },
        "date": 1754525337172,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 252092,
            "range": "±1.04%",
            "unit": "ops/sec",
            "extra": "231 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 243427,
            "range": "±0.83%",
            "unit": "ops/sec",
            "extra": "231 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "137432604+Ugonnaak1@users.noreply.github.com",
            "name": "Ugonna Akali",
            "username": "Ugonnaak1"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "dc60b43a3a5619f510da62d02be48bcddd3aaf5a",
          "message": "Surface Errors from MsalRuntime with Interaction Required (#7961)\n\nWhen customers use `acquireTokenSilent` with msal-node-runtime, errors\nreported to OneAuth-MSAL (such as interaction_required) are surfaced as:\n\n```\n  \"errormessage\": \"interaction_required: (pii)\",\n  \"errorname\": \"InteractionRequiredAuthError\",\n  \"errorstack\": \"InteractionRequiredAuthError: interaction_required: (pii)\\n    at ue.wrapError (c:\\\\Program <REDACTED: user-file-path> VS <REDACTED: user-file-path>:2:386244)\\n    at Object.o (c:\\\\Program <REDACTED: user-file-path> VS <REDACTED: user-file-path>:2:381487)\"\n```\n \nHowever, these errors lack critical context such as the error code and\nerror tag from the broker or our library which makes it difficult for\nour team to diagnose and resolve their issues.\n\n\nThis PR helps surface errors for interaction-required scenarios by \nreplacing `InteractionRequiredAuthError` with `NativeAuthError` for\nInteractionRequired Status in NativeBrokerPlugin and enhancing the\nNativeAuthError context\n\n---------\n\nCo-authored-by: Copilot <175728472+Copilot@users.noreply.github.com>",
          "timestamp": "2025-08-06T15:48:24-07:00",
          "tree_id": "bc959580459691b12c9d16697e031bc9e8d52076",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/dc60b43a3a5619f510da62d02be48bcddd3aaf5a"
        },
        "date": 1754540846430,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 252494,
            "range": "±0.77%",
            "unit": "ops/sec",
            "extra": "234 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 251700,
            "range": "±0.73%",
            "unit": "ops/sec",
            "extra": "224 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "kshabelko@microsoft.com",
            "name": "Konstantin",
            "username": "konstantin-msft"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "1c05d6c980d950ef3111ca369c161eecd09b08b9",
          "message": "Instrument timed out or cancelled pre-redirect requests (#7984)\n\n- Instrument timed out or cancelled pre-redirect requests",
          "timestamp": "2025-08-08T15:12:26-04:00",
          "tree_id": "1149e8d2d40a405543c10340467cdd2ecb60f894",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/1c05d6c980d950ef3111ca369c161eecd09b08b9"
        },
        "date": 1754691246558,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 247644,
            "range": "±0.87%",
            "unit": "ops/sec",
            "extra": "236 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 244078,
            "range": "±0.71%",
            "unit": "ops/sec",
            "extra": "212 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "137432604+Ugonnaak1@users.noreply.github.com",
            "name": "Ugonna Akali",
            "username": "Ugonnaak1"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "ebcfec14bc4f20e5dad40ba38d491492948cefee",
          "message": "update msal-node-runtime version (#7979)",
          "timestamp": "2025-08-12T07:21:46-07:00",
          "tree_id": "8b9251b9edf6b5d0d15400b719ad7701eaf4ff89",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/ebcfec14bc4f20e5dad40ba38d491492948cefee"
        },
        "date": 1755011675305,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 250506,
            "range": "±0.84%",
            "unit": "ops/sec",
            "extra": "236 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 247605,
            "range": "±0.78%",
            "unit": "ops/sec",
            "extra": "233 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "152663010+msal-js-release-automation[bot]@users.noreply.github.com",
            "name": "msal-js-release-automation[bot]",
            "username": "msal-js-release-automation[bot]"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "2e357d620c1c9425e54286ecaac1fb4839d83fa5",
          "message": "Release PR: official (#7994)\n\nThis PR contains the changelogs and version bumps for the MSAL.js 3P\nreleases.\n\nCo-authored-by: MSAL.js Release Automation <msaljsbuilds@microsoft.com>",
          "timestamp": "2025-08-12T20:38:35-07:00",
          "tree_id": "829d23e74dc718d9a0b6227be3067db40613f21c",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/2e357d620c1c9425e54286ecaac1fb4839d83fa5"
        },
        "date": 1755077639396,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 249122,
            "range": "±0.90%",
            "unit": "ops/sec",
            "extra": "234 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 250801,
            "range": "±0.72%",
            "unit": "ops/sec",
            "extra": "233 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "thomas.norling@microsoft.com",
            "name": "Thomas Norling",
            "username": "tnorling"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "f054cf608d2544858d8a3dbd02efd8b464ca862a",
          "message": "Update CODEOWNERS with MSAL.js Team (#7996)\n\nThis pull request updates the `CODEOWNERS` file to consolidate and\nsimplify code ownership assignments. The main change is replacing\nindividual maintainer usernames with the\n`@AzureAD/msal-js-public-client` team for most areas, and updating\nnode-related paths to include both `@AzureAD/msal-js-public-client` and\n`@AzureAD/id4s-msal-team`. This streamlines code review responsibilities\nand reflects current team structures.\n\nOwnership assignment updates:\n\n* Replaced individual user assignments with the\n`@AzureAD/msal-js-public-client` team for general ownership and for\nseveral library and sample directories.\n* Updated ownership for `/lib/msal-node/` and\n`/samples/msal-node-samples/` to include both\n`@AzureAD/msal-js-public-client` and `@AzureAD/id4s-msal-team`, ensuring\nboth teams are responsible for these areas.\n\nCleanup and simplification:\n\n* Removed redundant or outdated individual user assignments from various\ndirectories, consolidating ownership under relevant teams.",
          "timestamp": "2025-08-14T17:16:03-07:00",
          "tree_id": "9a577f671644785be69598636274eeccb4131df8",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/f054cf608d2544858d8a3dbd02efd8b464ca862a"
        },
        "date": 1755241139110,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 256037,
            "range": "±0.76%",
            "unit": "ops/sec",
            "extra": "236 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 254167,
            "range": "±0.71%",
            "unit": "ops/sec",
            "extra": "235 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "kshabelko@microsoft.com",
            "name": "Konstantin",
            "username": "konstantin-msft"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "e50a9c142ac7ef147eac0119619bbab7b235bbc4",
          "message": "Add bundle minification practices to copilot instructions (#7963)\n\n- Add bundle minification practices to copilot instructions",
          "timestamp": "2025-08-15T06:37:42-07:00",
          "tree_id": "63de82085254668d8cb1e4e7ab2100590b06796e",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/e50a9c142ac7ef147eac0119619bbab7b235bbc4"
        },
        "date": 1755271645425,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 245573,
            "range": "±0.86%",
            "unit": "ops/sec",
            "extra": "233 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 251155,
            "range": "±0.79%",
            "unit": "ops/sec",
            "extra": "235 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "198982749+Copilot@users.noreply.github.com",
            "name": "Copilot",
            "username": "Copilot"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "7fe97cf91a66df288edee99624e1dc5d32496977",
          "message": "Fix redirect loop when URLs contain encoded apostrophes in MSAL Angular standalone components (#7878)\n\n## Problem\n\nWhen using MSAL Angular standalone components, users experience infinite\nredirect loops after authentication when the URL contains encoded\napostrophes (`%27`) in query parameters. For example:\n\n```\nhttps://localhost:4200/profile?comments=blah%27blah\n```\n\nAfter authentication, the app gets stuck in a redirect loop instead of\ndisplaying the intended page.\n\n## Root Cause\n\nThe issue occurs in `RedirectClient.handleRedirectPromise()` during URL\ncomparison. The method compares the stored login request URL with the\ncurrent URL to determine if navigation is needed. However, the\ncomparison doesn't handle URL encoding consistently:\n\n- **Stored URL**: `https://localhost:4200/profile?comments=blah%27blah`\n(encoded apostrophe)\n- **Current URL**: `https://localhost:4200/profile?comments=blah'blah`\n(decoded apostrophe)\n\nSince `%27` ≠ `'` after normalization, MSAL thinks it's not on the\ncorrect page and attempts to navigate back, causing an infinite loop.\n\n## Solution\n\nAdded a new `normalizeUrlForComparison()` method in `RedirectClient`\nthat:\n\n1. Uses the native `URL` constructor to handle encoding consistently\n2. Ensures both URLs are normalized to the same encoding format\n3. Preserves existing canonicalization logic\n4. Includes graceful error handling with fallback\n\n```typescript\nprivate normalizeUrlForComparison(url: string): string {\n    if (!url) return url;\n    \n    const urlWithoutHash = url.split(\"#\")[0];\n    try {\n        const urlObj = new URL(urlWithoutHash);\n        const normalizedUrl = urlObj.origin + urlObj.pathname + urlObj.search;\n        return UrlString.canonicalizeUri(normalizedUrl);\n    } catch (e) {\n        // Fallback to original logic\n        return UrlString.canonicalizeUri(urlWithoutHash);\n    }\n}\n```\n\n## Testing\n\nAdded comprehensive test case covering:\n- ✅ Encoded vs decoded apostrophe scenario (the original issue)\n- ✅ Multiple encoded characters\n- ✅ Hash handling in redirect scenarios\n- ✅ Edge cases and error conditions\n\n## Impact\n\n- **Fixes redirect loops** for URLs with encoded special characters\n- **Zero breaking changes** - maintains backward compatibility\n- **Minimal performance impact** - only affects URL comparison logic\n- **Robust solution** - handles all URL-encoded characters consistently\n\n## Before/After\n\n**Before (broken):**\n```\nStored:  https://localhost:4200/profile?comments=blah%27blah\nCurrent: https://localhost:4200/profile?comments=blah'blah\nMatch: false → Redirect loop\n```\n\n**After (fixed):**\n```  \nStored:  https://localhost:4200/profile?comments=blah%27blah\nCurrent: https://localhost:4200/profile?comments=blah'blah  \nMatch: true → Normal flow continues\n```\n\nFixes #7636.\n\n> [!WARNING]\n>\n> <details>\n> <summary>Firewall rules blocked me from connecting to one or more\naddresses</summary>\n>\n> #### I tried to connect to the following addresses, but was blocked by\nfirewall rules:\n>\n> - `googlechromelabs.github.io`\n>   - Triggering command: `node install.mjs ` (dns block)\n> -\n`https://storage.googleapis.com/chrome-for-testing-public/132.0.6834.110/linux64/chrome-linux64.zip`\n>   - Triggering command: `node install.mjs ` (http block)\n>\n> If you need me to access, download, or install something from one of\nthese locations, you can either:\n>\n> - Configure [Actions setup\nsteps](https://gh.io/copilot/actions-setup-steps) to set up my\nenvironment, which run before the firewall is enabled\n> - Add the appropriate URLs or hosts to my [firewall allow\nlist](https://gh.io/copilot/firewall-config)\n>\n> </details>\n\n\n\n<!-- START COPILOT CODING AGENT TIPS -->\n---\n\n💬 Share your feedback on Copilot coding agent for the chance to win a\n$200 gift card! Click\n[here](https://survey.alchemer.com/s3/8343779/Copilot-Coding-agent) to\nstart the survey.\n\n---------\n\nCo-authored-by: copilot-swe-agent[bot] <198982749+Copilot@users.noreply.github.com>\nCo-authored-by: tnorling <5307810+tnorling@users.noreply.github.com>\nCo-authored-by: Thomas Norling <thomas.norling@microsoft.com>",
          "timestamp": "2025-08-15T10:47:31-07:00",
          "tree_id": "40cf3753e0c23a798ec0a67d5ef80b6e147d47b6",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/7fe97cf91a66df288edee99624e1dc5d32496977"
        },
        "date": 1755293577508,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 252650,
            "range": "±0.90%",
            "unit": "ops/sec",
            "extra": "234 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 254889,
            "range": "±0.79%",
            "unit": "ops/sec",
            "extra": "235 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "198982749+Copilot@users.noreply.github.com",
            "name": "Copilot",
            "username": "Copilot"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "81f41cb452dd85e1107036c226dbb0fa08caafaf",
          "message": "Fix cache not used for getting token if scopes are empty (#7995)\n\nThe issue occurs when `acquireTokenSilent` is called with empty scopes\n(`scopes: []`). Instead of using cached tokens, the library throws a\n`ClientConfigurationError` during cache lookup and makes unnecessary API\nrequests to Azure AD.\n\n## Root Cause\nThe problem is in `ScopeSet.createSearchScopes()` which is called from\n`CacheManager.getAccessToken()` during cache lookup. When empty scopes\nare passed, the `ScopeSet` constructor throws an error because it\ndoesn't allow empty scope arrays, preventing cache lookup from\ncompleting.\n\n## Solution\nModified `ScopeSet.createSearchScopes()` to handle empty, null, or\nundefined scopes by providing default OIDC scopes (`openid`, `profile`,\n`offline_access`) before calling the constructor. This approach:\n\n- Follows the same pattern as `RequestParameterBuilder.addScopes()`\nwhich already handles empty scopes\n- Allows cache lookup to proceed with reasonable default scopes when no\nspecific scopes are requested\n- Maintains all existing behavior for non-empty scopes\n- Eliminates unnecessary network requests when tokens are already cached\n\n## Example\n```javascript\nconst { instance, accounts } = useMsal();\nconst account = useAccount(accounts[0]);\n\n// This now works and uses cache instead of making API calls\nconst response = await instance.acquireTokenSilent({\n    scopes: [], // Empty scopes now supported\n    account\n});\n```\n\nThe fix enables `acquireTokenSilent` to properly utilize cached tokens\nwhen called with empty scopes, improving performance and user\nexperience.\n\nFixes #6969.\n\n<!-- START COPILOT CODING AGENT TIPS -->\n---\n\n💬 Share your feedback on Copilot coding agent for the chance to win a\n$200 gift card! Click\n[here](https://survey.alchemer.com/s3/8343779/Copilot-Coding-agent) to\nstart the survey.\n\n---------\n\nCo-authored-by: copilot-swe-agent[bot] <198982749+Copilot@users.noreply.github.com>\nCo-authored-by: tnorling <5307810+tnorling@users.noreply.github.com>\nCo-authored-by: Thomas Norling <thomas.norling@microsoft.com>\nCo-authored-by: Copilot <175728472+Copilot@users.noreply.github.com>",
          "timestamp": "2025-08-15T14:48:04-07:00",
          "tree_id": "044e5d8d5a56afd50221f7c681f742d18d1d43f3",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/81f41cb452dd85e1107036c226dbb0fa08caafaf"
        },
        "date": 1755316498622,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 254096,
            "range": "±0.70%",
            "unit": "ops/sec",
            "extra": "235 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 251655,
            "range": "±0.80%",
            "unit": "ops/sec",
            "extra": "236 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "sameera.gajjarapu@microsoft.com",
            "name": "Sameera Gajjarapu",
            "username": "sameerag"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "866f2da124aa13700cfe9d14d5a78ddda9d2bf53",
          "message": "Add JS platform telemetry params (#7991)\n\n- `isPlatformAuthorizeRequest:boolean` Is set on every request that is\nsent to STS with nativeBroker=1\n- `isPlatformBrokerRequest:boolean` Is set on every request that is sent\nto the platform broker directly, and always set only if\n`nativeAccountId` is in the cache/request\n- `isNativeBroker:boolean` Is set on every successful response from the\nBroker\n- `BrokerErrorName` for intermittent fatal broker errors\n\n---------\n\nCo-authored-by: Copilot <175728472+Copilot@users.noreply.github.com>",
          "timestamp": "2025-08-18T14:50:34-07:00",
          "tree_id": "ebfd699a9928d97e95ef27dd84025c86e0d5d919",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/866f2da124aa13700cfe9d14d5a78ddda9d2bf53"
        },
        "date": 1755573276958,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 249940,
            "range": "±0.83%",
            "unit": "ops/sec",
            "extra": "235 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 243924,
            "range": "±0.76%",
            "unit": "ops/sec",
            "extra": "232 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "152663010+msal-js-release-automation[bot]@users.noreply.github.com",
            "name": "msal-js-release-automation[bot]",
            "username": "msal-js-release-automation[bot]"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "d8718bc1ff1a5ebc8d2bb5138ddaeb72cdd680a6",
          "message": "Release PR: official (#8008)\n\nThis PR contains the changelogs and version bumps for the MSAL.js 3P\nreleases.\n\nCo-authored-by: MSAL.js Release Automation <msaljsbuilds@microsoft.com>",
          "timestamp": "2025-08-19T16:53:35-07:00",
          "tree_id": "eb6c961f79b65c34af5c6991792cb4a7068cfa43",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/d8718bc1ff1a5ebc8d2bb5138ddaeb72cdd680a6"
        },
        "date": 1755659775110,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 242812,
            "range": "±0.83%",
            "unit": "ops/sec",
            "extra": "223 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 242935,
            "range": "±0.88%",
            "unit": "ops/sec",
            "extra": "209 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "thomas.norling@microsoft.com",
            "name": "Thomas Norling",
            "username": "tnorling"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "190730d86116ba71bc407c194cc7c9d1fd1c94ef",
          "message": "Fix CODEOWNERS for custom-auth (#8011)\n\nFixes the Codeowners path for the custom-auth features",
          "timestamp": "2025-08-21T18:07:22Z",
          "tree_id": "5ff7125c274f36009d045ae2db7aea061d828e99",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/190730d86116ba71bc407c194cc7c9d1fd1c94ef"
        },
        "date": 1755810882415,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 233846,
            "range": "±0.93%",
            "unit": "ops/sec",
            "extra": "233 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 237560,
            "range": "±0.96%",
            "unit": "ops/sec",
            "extra": "233 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "thomas.norling@microsoft.com",
            "name": "Thomas Norling",
            "username": "tnorling"
          },
          "committer": {
            "email": "kshabelko@microsoft.com",
            "name": "Konstantin",
            "username": "konstantin-msft"
          },
          "distinct": true,
          "id": "2e533d593544f82bbb01457cf1bda403e6d7f708",
          "message": "resolve comments",
          "timestamp": "2025-08-22T10:33:46-04:00",
          "tree_id": "1b8fe5525d4d71bb47e0cd02e17d9d27641d8eeb",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/2e533d593544f82bbb01457cf1bda403e6d7f708"
        },
        "date": 1755875683548,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 241964,
            "range": "±0.89%",
            "unit": "ops/sec",
            "extra": "228 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 242054,
            "range": "±0.70%",
            "unit": "ops/sec",
            "extra": "233 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "49699333+dependabot[bot]@users.noreply.github.com",
            "name": "dependabot[bot]",
            "username": "dependabot[bot]"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "d38ec1b8ea84fcc660f655560dbc518c90e374e1",
          "message": "Bump actions/checkout from 4 to 5 (#8003)\n\nBumps [actions/checkout](https://github.com/actions/checkout) from 4 to\n5.\n<details>\n<summary>Release notes</summary>\n<p><em>Sourced from <a\nhref=\"https://github.com/actions/checkout/releases\">actions/checkout's\nreleases</a>.</em></p>\n<blockquote>\n<h2>v5.0.0</h2>\n<h2>What's Changed</h2>\n<ul>\n<li>Update actions checkout to use node 24 by <a\nhref=\"https://github.com/salmanmkc\"><code>@​salmanmkc</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/2226\">actions/checkout#2226</a></li>\n<li>Prepare v5.0.0 release by <a\nhref=\"https://github.com/salmanmkc\"><code>@​salmanmkc</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/2238\">actions/checkout#2238</a></li>\n</ul>\n<h2>⚠️ Minimum Compatible Runner Version</h2>\n<p><strong>v2.327.1</strong><br />\n<a\nhref=\"https://github.com/actions/runner/releases/tag/v2.327.1\">Release\nNotes</a></p>\n<p>Make sure your runner is updated to this version or newer to use this\nrelease.</p>\n<p><strong>Full Changelog</strong>: <a\nhref=\"https://github.com/actions/checkout/compare/v4...v5.0.0\">https://github.com/actions/checkout/compare/v4...v5.0.0</a></p>\n<h2>v4.3.0</h2>\n<h2>What's Changed</h2>\n<ul>\n<li>docs: update README.md by <a\nhref=\"https://github.com/motss\"><code>@​motss</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1971\">actions/checkout#1971</a></li>\n<li>Add internal repos for checking out multiple repositories by <a\nhref=\"https://github.com/mouismail\"><code>@​mouismail</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1977\">actions/checkout#1977</a></li>\n<li>Documentation update - add recommended permissions to Readme by <a\nhref=\"https://github.com/benwells\"><code>@​benwells</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/2043\">actions/checkout#2043</a></li>\n<li>Adjust positioning of user email note and permissions heading by <a\nhref=\"https://github.com/joshmgross\"><code>@​joshmgross</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/2044\">actions/checkout#2044</a></li>\n<li>Update README.md by <a\nhref=\"https://github.com/nebuk89\"><code>@​nebuk89</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/2194\">actions/checkout#2194</a></li>\n<li>Update CODEOWNERS for actions by <a\nhref=\"https://github.com/TingluoHuang\"><code>@​TingluoHuang</code></a>\nin <a\nhref=\"https://redirect.github.com/actions/checkout/pull/2224\">actions/checkout#2224</a></li>\n<li>Update package dependencies by <a\nhref=\"https://github.com/salmanmkc\"><code>@​salmanmkc</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/2236\">actions/checkout#2236</a></li>\n<li>Prepare release v4.3.0 by <a\nhref=\"https://github.com/salmanmkc\"><code>@​salmanmkc</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/2237\">actions/checkout#2237</a></li>\n</ul>\n<h2>New Contributors</h2>\n<ul>\n<li><a href=\"https://github.com/motss\"><code>@​motss</code></a> made\ntheir first contribution in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1971\">actions/checkout#1971</a></li>\n<li><a href=\"https://github.com/mouismail\"><code>@​mouismail</code></a>\nmade their first contribution in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1977\">actions/checkout#1977</a></li>\n<li><a href=\"https://github.com/benwells\"><code>@​benwells</code></a>\nmade their first contribution in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/2043\">actions/checkout#2043</a></li>\n<li><a href=\"https://github.com/nebuk89\"><code>@​nebuk89</code></a> made\ntheir first contribution in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/2194\">actions/checkout#2194</a></li>\n<li><a href=\"https://github.com/salmanmkc\"><code>@​salmanmkc</code></a>\nmade their first contribution in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/2236\">actions/checkout#2236</a></li>\n</ul>\n<p><strong>Full Changelog</strong>: <a\nhref=\"https://github.com/actions/checkout/compare/v4...v4.3.0\">https://github.com/actions/checkout/compare/v4...v4.3.0</a></p>\n<h2>v4.2.2</h2>\n<h2>What's Changed</h2>\n<ul>\n<li><code>url-helper.ts</code> now leverages well-known environment\nvariables by <a href=\"https://github.com/jww3\"><code>@​jww3</code></a>\nin <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1941\">actions/checkout#1941</a></li>\n<li>Expand unit test coverage for <code>isGhes</code> by <a\nhref=\"https://github.com/jww3\"><code>@​jww3</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1946\">actions/checkout#1946</a></li>\n</ul>\n<p><strong>Full Changelog</strong>: <a\nhref=\"https://github.com/actions/checkout/compare/v4.2.1...v4.2.2\">https://github.com/actions/checkout/compare/v4.2.1...v4.2.2</a></p>\n<h2>v4.2.1</h2>\n<h2>What's Changed</h2>\n<ul>\n<li>Check out other refs/* by commit if provided, fall back to ref by <a\nhref=\"https://github.com/orhantoy\"><code>@​orhantoy</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1924\">actions/checkout#1924</a></li>\n</ul>\n<h2>New Contributors</h2>\n<ul>\n<li><a href=\"https://github.com/Jcambass\"><code>@​Jcambass</code></a>\nmade their first contribution in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1919\">actions/checkout#1919</a></li>\n</ul>\n<p><strong>Full Changelog</strong>: <a\nhref=\"https://github.com/actions/checkout/compare/v4.2.0...v4.2.1\">https://github.com/actions/checkout/compare/v4.2.0...v4.2.1</a></p>\n<!-- raw HTML omitted -->\n</blockquote>\n<p>... (truncated)</p>\n</details>\n<details>\n<summary>Changelog</summary>\n<p><em>Sourced from <a\nhref=\"https://github.com/actions/checkout/blob/main/CHANGELOG.md\">actions/checkout's\nchangelog</a>.</em></p>\n<blockquote>\n<h1>Changelog</h1>\n<h2>V5.0.0</h2>\n<ul>\n<li>Update actions checkout to use node 24 by <a\nhref=\"https://github.com/salmanmkc\"><code>@​salmanmkc</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/2226\">actions/checkout#2226</a></li>\n</ul>\n<h2>V4.3.0</h2>\n<ul>\n<li>docs: update README.md by <a\nhref=\"https://github.com/motss\"><code>@​motss</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1971\">actions/checkout#1971</a></li>\n<li>Add internal repos for checking out multiple repositories by <a\nhref=\"https://github.com/mouismail\"><code>@​mouismail</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1977\">actions/checkout#1977</a></li>\n<li>Documentation update - add recommended permissions to Readme by <a\nhref=\"https://github.com/benwells\"><code>@​benwells</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/2043\">actions/checkout#2043</a></li>\n<li>Adjust positioning of user email note and permissions heading by <a\nhref=\"https://github.com/joshmgross\"><code>@​joshmgross</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/2044\">actions/checkout#2044</a></li>\n<li>Update README.md by <a\nhref=\"https://github.com/nebuk89\"><code>@​nebuk89</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/2194\">actions/checkout#2194</a></li>\n<li>Update CODEOWNERS for actions by <a\nhref=\"https://github.com/TingluoHuang\"><code>@​TingluoHuang</code></a>\nin <a\nhref=\"https://redirect.github.com/actions/checkout/pull/2224\">actions/checkout#2224</a></li>\n<li>Update package dependencies by <a\nhref=\"https://github.com/salmanmkc\"><code>@​salmanmkc</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/2236\">actions/checkout#2236</a></li>\n</ul>\n<h2>v4.2.2</h2>\n<ul>\n<li><code>url-helper.ts</code> now leverages well-known environment\nvariables by <a href=\"https://github.com/jww3\"><code>@​jww3</code></a>\nin <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1941\">actions/checkout#1941</a></li>\n<li>Expand unit test coverage for <code>isGhes</code> by <a\nhref=\"https://github.com/jww3\"><code>@​jww3</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1946\">actions/checkout#1946</a></li>\n</ul>\n<h2>v4.2.1</h2>\n<ul>\n<li>Check out other refs/* by commit if provided, fall back to ref by <a\nhref=\"https://github.com/orhantoy\"><code>@​orhantoy</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1924\">actions/checkout#1924</a></li>\n</ul>\n<h2>v4.2.0</h2>\n<ul>\n<li>Add Ref and Commit outputs by <a\nhref=\"https://github.com/lucacome\"><code>@​lucacome</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1180\">actions/checkout#1180</a></li>\n<li>Dependency updates by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>- <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1777\">actions/checkout#1777</a>,\n<a\nhref=\"https://redirect.github.com/actions/checkout/pull/1872\">actions/checkout#1872</a></li>\n</ul>\n<h2>v4.1.7</h2>\n<ul>\n<li>Bump the minor-npm-dependencies group across 1 directory with 4\nupdates by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1739\">actions/checkout#1739</a></li>\n<li>Bump actions/checkout from 3 to 4 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1697\">actions/checkout#1697</a></li>\n<li>Check out other refs/* by commit by <a\nhref=\"https://github.com/orhantoy\"><code>@​orhantoy</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1774\">actions/checkout#1774</a></li>\n<li>Pin actions/checkout's own workflows to a known, good, stable\nversion. by <a href=\"https://github.com/jww3\"><code>@​jww3</code></a> in\n<a\nhref=\"https://redirect.github.com/actions/checkout/pull/1776\">actions/checkout#1776</a></li>\n</ul>\n<h2>v4.1.6</h2>\n<ul>\n<li>Check platform to set archive extension appropriately by <a\nhref=\"https://github.com/cory-miller\"><code>@​cory-miller</code></a> in\n<a\nhref=\"https://redirect.github.com/actions/checkout/pull/1732\">actions/checkout#1732</a></li>\n</ul>\n<h2>v4.1.5</h2>\n<ul>\n<li>Update NPM dependencies by <a\nhref=\"https://github.com/cory-miller\"><code>@​cory-miller</code></a> in\n<a\nhref=\"https://redirect.github.com/actions/checkout/pull/1703\">actions/checkout#1703</a></li>\n<li>Bump github/codeql-action from 2 to 3 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1694\">actions/checkout#1694</a></li>\n<li>Bump actions/setup-node from 1 to 4 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1696\">actions/checkout#1696</a></li>\n<li>Bump actions/upload-artifact from 2 to 4 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1695\">actions/checkout#1695</a></li>\n<li>README: Suggest <code>user.email</code> to be\n<code>41898282+github-actions[bot]@users.noreply.github.com</code> by <a\nhref=\"https://github.com/cory-miller\"><code>@​cory-miller</code></a> in\n<a\nhref=\"https://redirect.github.com/actions/checkout/pull/1707\">actions/checkout#1707</a></li>\n</ul>\n<h2>v4.1.4</h2>\n<ul>\n<li>Disable <code>extensions.worktreeConfig</code> when disabling\n<code>sparse-checkout</code> by <a\nhref=\"https://github.com/jww3\"><code>@​jww3</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1692\">actions/checkout#1692</a></li>\n<li>Add dependabot config by <a\nhref=\"https://github.com/cory-miller\"><code>@​cory-miller</code></a> in\n<a\nhref=\"https://redirect.github.com/actions/checkout/pull/1688\">actions/checkout#1688</a></li>\n<li>Bump the minor-actions-dependencies group with 2 updates by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1693\">actions/checkout#1693</a></li>\n<li>Bump word-wrap from 1.2.3 to 1.2.5 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a> in <a\nhref=\"https://redirect.github.com/actions/checkout/pull/1643\">actions/checkout#1643</a></li>\n</ul>\n<h2>v4.1.3</h2>\n<!-- raw HTML omitted -->\n</blockquote>\n<p>... (truncated)</p>\n</details>\n<details>\n<summary>Commits</summary>\n<ul>\n<li><a\nhref=\"https://github.com/actions/checkout/commit/08c6903cd8c0fde910a37f88322edcfb5dd907a8\"><code>08c6903</code></a>\nPrepare v5.0.0 release (<a\nhref=\"https://redirect.github.com/actions/checkout/issues/2238\">#2238</a>)</li>\n<li><a\nhref=\"https://github.com/actions/checkout/commit/9f265659d3bb64ab1440b03b12f4d47a24320917\"><code>9f26565</code></a>\nUpdate actions checkout to use node 24 (<a\nhref=\"https://redirect.github.com/actions/checkout/issues/2226\">#2226</a>)</li>\n<li>See full diff in <a\nhref=\"https://github.com/actions/checkout/compare/v4...v5\">compare\nview</a></li>\n</ul>\n</details>\n<br />\n\n\n[![Dependabot compatibility\nscore](https://dependabot-badges.githubapp.com/badges/compatibility_score?dependency-name=actions/checkout&package-manager=github_actions&previous-version=4&new-version=5)](https://docs.github.com/en/github/managing-security-vulnerabilities/about-dependabot-security-updates#about-compatibility-scores)\n\nDependabot will resolve any conflicts with this PR as long as you don't\nalter it yourself. You can also trigger a rebase manually by commenting\n`@dependabot rebase`.\n\n[//]: # (dependabot-automerge-start)\n[//]: # (dependabot-automerge-end)\n\n---\n\n<details>\n<summary>Dependabot commands and options</summary>\n<br />\n\nYou can trigger Dependabot actions by commenting on this PR:\n- `@dependabot rebase` will rebase this PR\n- `@dependabot recreate` will recreate this PR, overwriting any edits\nthat have been made to it\n- `@dependabot merge` will merge this PR after your CI passes on it\n- `@dependabot squash and merge` will squash and merge this PR after\nyour CI passes on it\n- `@dependabot cancel merge` will cancel a previously requested merge\nand block automerging\n- `@dependabot reopen` will reopen this PR if it is closed\n- `@dependabot close` will close this PR and stop Dependabot recreating\nit. You can achieve the same result by closing it manually\n- `@dependabot show <dependency name> ignore conditions` will show all\nof the ignore conditions of the specified dependency\n- `@dependabot ignore this major version` will close this PR and stop\nDependabot creating any more for this major version (unless you reopen\nthe PR or upgrade to it yourself)\n- `@dependabot ignore this minor version` will close this PR and stop\nDependabot creating any more for this minor version (unless you reopen\nthe PR or upgrade to it yourself)\n- `@dependabot ignore this dependency` will close this PR and stop\nDependabot creating any more for this dependency (unless you reopen the\nPR or upgrade to it yourself)\n\n\n</details>\n\nSigned-off-by: dependabot[bot] <support@github.com>\nCo-authored-by: dependabot[bot] <49699333+dependabot[bot]@users.noreply.github.com>\nCo-authored-by: Thomas Norling <thomas.norling@microsoft.com>",
          "timestamp": "2025-08-22T19:03:03-04:00",
          "tree_id": "177e3708cfcfdd3da6358599b72eb5a3057b9491",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/d38ec1b8ea84fcc660f655560dbc518c90e374e1"
        },
        "date": 1755918874003,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 240241,
            "range": "±0.75%",
            "unit": "ops/sec",
            "extra": "235 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 238736,
            "range": "±0.83%",
            "unit": "ops/sec",
            "extra": "233 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "thomas.norling@microsoft.com",
            "name": "Thomas Norling",
            "username": "tnorling"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "d1cfe74648e216fcdeb9d915eb3df8f41085bfc4",
          "message": "Fix Custom-Auth Codeowners (#8015)\n\nFilepath was incorrect, fixed.",
          "timestamp": "2025-08-25T11:29:35-07:00",
          "tree_id": "293d911f15064b649cc0ea8641911bb37b802ac8",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/d1cfe74648e216fcdeb9d915eb3df8f41085bfc4"
        },
        "date": 1756157285462,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 245005,
            "range": "±0.76%",
            "unit": "ops/sec",
            "extra": "236 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 246188,
            "range": "±0.70%",
            "unit": "ops/sec",
            "extra": "235 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "137432604+Ugonnaak1@users.noreply.github.com",
            "name": "Ugonna Akali",
            "username": "Ugonnaak1"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "1a521c3e7631d4ed50d3082d26036162a4ef483e",
          "message": "Enable passing of redirect uri for Nativebroker (#8005)\n\nCurrent PCA requests don't allow the customer to pass in their own\nredirecturi. This is an issue for signed MacOS apps because the customer\nis forced to use the unsigned redirect URI which will hit a redirecturi\nvalidation error in the Apple Broker\n```\n  errorMessage: 'Description: Error Domain=MSALErrorDomain Code=-50000 \"(null)\" UserInfo={MSALErrorDescriptionKey=MSAL redirectUri validation error: redirect uri has incorrect scheme - it must be in the form of msauth.<app_bundle_id>://auth\\n' +\n    'ADAL redirectUri validation error: Source application does not match redirect uri host. Invalid source app., MSALInternalErrorCodeKey=-42000, MSALBrokerVersionKey=5.2508.0}, Domain: MSALErrorDomain.Error was thrown in sourceArea: Broker (Error Code: -42000, Tag: 508638916)',\n\n```\nSolution:\nMake providing a redirecturi optional. If the customer doesn't provide\none, we fallback to the default for each platform in broker scenarios\nFor non-broker scenarios, if redirecturi is provided, we warn them it\nwon't be used and use the default\n\nAlso added a function to convert error tags to their string version to\nmake error lookup quicker.",
          "timestamp": "2025-08-25T16:02:44-07:00",
          "tree_id": "fd69ef2495af9050418168f7c14fce9fee97c43b",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/1a521c3e7631d4ed50d3082d26036162a4ef483e"
        },
        "date": 1756178082034,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 256613,
            "range": "±0.91%",
            "unit": "ops/sec",
            "extra": "232 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 259213,
            "range": "±1.03%",
            "unit": "ops/sec",
            "extra": "235 samples"
          }
        ]
      }
    ]
  }
}