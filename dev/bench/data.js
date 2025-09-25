window.BENCHMARK_DATA = {
  "lastUpdate": 1758834882781,
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
          "id": "8316ee5f4bf29bdb326bf51b9b621b29731b36dc",
          "message": "Add extraQueryParams for api calls (#7974)",
          "timestamp": "2025-08-27T19:50:29+01:00",
          "tree_id": "f7501b1a73f3c1c1226d4fb37f378613125226ad",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/8316ee5f4bf29bdb326bf51b9b621b29731b36dc"
        },
        "date": 1756329267119,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 249792,
            "range": "±0.75%",
            "unit": "ops/sec",
            "extra": "212 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 245564,
            "range": "±0.92%",
            "unit": "ops/sec",
            "extra": "224 samples"
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
          "id": "6d975bcac88832b21e8e1ac030d3802133640005",
          "message": "Upgrade/Rollback E2E Tests (#8017)\n\nThis pull request introduces a new end-to-end (E2E) test setup for the\n`ExpressSample` in the MSAL browser samples, along with several\nimprovements to the sample app's authentication flow, version switching,\nand UI. The most significant changes are the addition of E2E test\ninfrastructure, enhancements to how authentication and profile data are\nhandled and displayed, and improvements to the version switching logic\nand UI.\n\n**E2E Test Infrastructure and Configuration:**\n\n* Added E2E test configuration and helpers for `ExpressSample`,\nincluding a new `.env.e2e` file, a `jest.config.cjs` configuration, and\na `test-helpers.ts` utility for Puppeteer-based tests. This enables\nautomated browser testing of authentication scenarios.\n[[1]](diffhunk://#diff-d257013e504cdeb26f2bfad2e809b36492f6fdbfc45ffcbf4bc55b2da4887c74R1-R5)\n[[2]](diffhunk://#diff-b70997a8822f049844c4b3afbe1ab62688a13f909a880256d0c8364fff169c3eR1-R8)\n[[3]](diffhunk://#diff-a45554c7aba0df12b0d4ff639ce1fe74f9e12a2904ecb1d0d2edbe002bc4eefeR1-R112)\n* Updated the pipeline configuration to include `ExpressSample` in E2E\ntest runs.\n* Added necessary devDependencies and scripts for E2E testing in\n`package.json`.\n\n**Authentication and Profile Data Handling:**\n\n* Modified the authentication flow to return the full MSAL\nauthentication response, not just the access token, and updated the MS\nGraph call logic to pass and display this data.\n[[1]](diffhunk://#diff-aedaa17dbc1e25e5997a95b1bea0c9c3b5a9caa6fbbc943c700972966e16b530L115-R115)\n[[2]](diffhunk://#diff-57185e771678ea1756556bb0a2d9f5cba1e5441609c75dc66cebf0106e92ce3bL11-R15)\n[[3]](diffhunk://#diff-57185e771678ea1756556bb0a2d9f5cba1e5441609c75dc66cebf0106e92ce3bL31-R30)\n[[4]](diffhunk://#diff-57185e771678ea1756556bb0a2d9f5cba1e5441609c75dc66cebf0106e92ce3bL49-R50)\n[[5]](diffhunk://#diff-57185e771678ea1756556bb0a2d9f5cba1e5441609c75dc66cebf0106e92ce3bR94-R117)\n* Added UI logic to display both profile data and raw authentication\ndata, with sensitive fields (like the access token) redacted for\ndisplay.\n\n**Version Switching and UI Improvements:**\n\n* Refactored version switching logic to use consistent keys (e.g.,\n`latest-v3` instead of `3.x`) and improved the UI for version selection,\nincluding unique IDs for dropdown items and better feedback when\nswitching versions.\n[[1]](diffhunk://#diff-7bf33d829c4c8178361e7badec48a371762d82042e849f0d2c94a2bb41001c8cL60-R57)\n[[2]](diffhunk://#diff-7bf33d829c4c8178361e7badec48a371762d82042e849f0d2c94a2bb41001c8cL70-R74)\n[[3]](diffhunk://#diff-7bf33d829c4c8178361e7badec48a371762d82042e849f0d2c94a2bb41001c8cL107-R104)\n[[4]](diffhunk://#diff-7bf33d829c4c8178361e7badec48a371762d82042e849f0d2c94a2bb41001c8cL129-R147)\n[[5]](diffhunk://#diff-2d7f351e21c92ea9ff545acdd7d5094e73a972fe350dea4df139f4e760a77d2eR92)\n[[6]](diffhunk://#diff-2d7f351e21c92ea9ff545acdd7d5094e73a972fe350dea4df139f4e760a77d2eR111)\n* Ensured UI elements (such as sign-in and version switcher buttons) are\nonly displayed after event handlers are registered, preventing flicker\nor premature interaction.\n[[1]](diffhunk://#diff-83705d1d56435a3dea7f60a3b8bfeb237387c901e34b0a7990233797b9ad2fbaR47-R48)\n[[2]](diffhunk://#diff-2d7f351e21c92ea9ff545acdd7d5094e73a972fe350dea4df139f4e760a77d2eR64-R66)\n\n**UI and Styling Enhancements:**\n\n* Updated CSS classes and selectors to use `.json-content` for\ndisplaying JSON data, improving clarity and maintainability.\n[[1]](diffhunk://#diff-0587667bae35b440adac264532b10d1e1c50fd1ad1cf003c42cbceb8365bac54L251-R252)\n[[2]](diffhunk://#diff-0587667bae35b440adac264532b10d1e1c50fd1ad1cf003c42cbceb8365bac54L260-R261)\n* Minor CSS and code style cleanups for readability and consistency.\n[[1]](diffhunk://#diff-0587667bae35b440adac264532b10d1e1c50fd1ad1cf003c42cbceb8365bac54L175-R176)\n[[2]](diffhunk://#diff-0587667bae35b440adac264532b10d1e1c50fd1ad1cf003c42cbceb8365bac54L523-R530)\n\n**Other Improvements:**\n\n* Removed unnecessary dependencies (like `morgan`) from the server and\ncleaned up middleware usage.\n[[1]](diffhunk://#diff-7bf33d829c4c8178361e7badec48a371762d82042e849f0d2c94a2bb41001c8cL8)\n[[2]](diffhunk://#diff-7bf33d829c4c8178361e7badec48a371762d82042e849f0d2c94a2bb41001c8cL39-L40)\n* Improved screenshot utility to always capture full-page screenshots\nduring E2E tests.\n\nThese changes collectively improve the testability, usability, and\nmaintainability of the `ExpressSample` app and its E2E testing setup.\n\n---------\n\nCo-authored-by: Copilot <175728472+Copilot@users.noreply.github.com>",
          "timestamp": "2025-08-28T13:59:16-07:00",
          "tree_id": "367ae66a96f94bc35add4705a603050c2201265a",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/6d975bcac88832b21e8e1ac030d3802133640005"
        },
        "date": 1756417303270,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 253231,
            "range": "±0.75%",
            "unit": "ops/sec",
            "extra": "226 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 252539,
            "range": "±0.77%",
            "unit": "ops/sec",
            "extra": "234 samples"
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
          "id": "5b86b0bdacc0b89a714b5a977dd19c63a80f4248",
          "message": "Bump next from 15.3.3 to 15.4.7 (#8023)\n\nBumps [next](https://github.com/vercel/next.js) from 15.3.3 to 15.4.7.\n<details>\n<summary>Release notes</summary>\n<p><em>Sourced from <a\nhref=\"https://github.com/vercel/next.js/releases\">next's\nreleases</a>.</em></p>\n<blockquote>\n<h2>v15.4.7</h2>\n<blockquote>\n<p>[!NOTE]<br />\nThis release is backporting bug fixes. It does <strong>not</strong>\ninclude all pending features/changes on canary.</p>\n</blockquote>\n<h3>Core Changes</h3>\n<ul>\n<li>fix router handling when setting a location response header <a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82588\">#82588</a></li>\n</ul>\n<h3>Credits</h3>\n<p>Huge thanks to <a\nhref=\"https://github.com/ztanner\"><code>@​ztanner</code></a> for\nhelping!</p>\n<h2>v15.4.6</h2>\n<blockquote>\n<p>[!NOTE]<br />\nThis release is backporting bug fixes. It does <strong>not</strong>\ninclude all pending features/changes on canary.</p>\n</blockquote>\n<h3>Core Changes</h3>\n<ul>\n<li>fix: <code>_error</code> page's <code>req.url</code> can be\noverwritten to dynamic param on minimal mode (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82347\">#82347</a>)</li>\n<li>fix: add <code>?dpl</code> to fonts in\n<code>/_next/static/media</code> (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82384\">#82384</a>)</li>\n</ul>\n<h3>Credits</h3>\n<p>Huge thanks to <a\nhref=\"https://github.com/devjiwonchoi\"><code>@​devjiwonchoi</code></a>,\n<a href=\"https://github.com/ijjk\"><code>@​ijjk</code></a>, and <a\nhref=\"https://github.com/styfle\"><code>@​styfle</code></a> for\nhelping!</p>\n<h2>v15.4.5</h2>\n<blockquote>\n<p>[!NOTE]<br />\nThis release is backporting bug fixes. It does <strong>not</strong>\ninclude all pending features/changes on canary.</p>\n</blockquote>\n<h3>Core Changes</h3>\n<ul>\n<li>Fix API stripping JSON incorrectly (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82062\">#82062</a>)</li>\n<li>Fix i18n fallback: false collision (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82158\">#82158</a>)</li>\n<li>Revert &quot;Fix tracing of server actions imported by client\ncomponents (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82167\">#82167</a>)</li>\n<li>Ensure setAssetPrefix updates config instance (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82165\">#82165</a>)</li>\n<li>Turbopack: update mimalloc (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82166\">#82166</a>)</li>\n<li>fix(next/image): fix image-optimizer.ts headers (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82175\">#82175</a>)</li>\n<li>fix(next/image): improve and simplify detect-content-type (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82174\">#82174</a>)</li>\n</ul>\n<h3>Credits</h3>\n<p>Huge thanks to <a\nhref=\"https://github.com/ijjk\"><code>@​ijjk</code></a>, <a\nhref=\"https://github.com/sokra\"><code>@​sokra</code></a>, and <a\nhref=\"https://github.com/styfle\"><code>@​styfle</code></a> for\nhelping!</p>\n<h2>v15.4.4</h2>\n<blockquote>\n<p>[!NOTE]<br />\nThis release is backporting bug fixes. It does <strong>not</strong>\ninclude all pending features/changes on canary.</p>\n</blockquote>\n<h3>Core Changes</h3>\n<ul>\n<li>Fix dynamicParams false layout case in dev (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82026\">#82026</a>)</li>\n<li>Turbopack: fix scope hoisting variable renaming bug (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/81640\">#81640</a>)</li>\n<li>Upgrade to swc v33 (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/81750\">#81750</a>)</li>\n<li>Revert &quot;[metadata] use https protocol for schema urls&quot; (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/81934\">#81934</a>)</li>\n</ul>\n<!-- raw HTML omitted -->\n</blockquote>\n<p>... (truncated)</p>\n</details>\n<details>\n<summary>Commits</summary>\n<ul>\n<li><a\nhref=\"https://github.com/vercel/next.js/commit/f30d815859e932e09222e93bb6e8a376b918d874\"><code>f30d815</code></a>\nv15.4.7</li>\n<li><a\nhref=\"https://github.com/vercel/next.js/commit/1a026e338d2b8c977c8949a134168952338d6d01\"><code>1a026e3</code></a>\nfix router handling when setting a location response header (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82588\">#82588</a>)</li>\n<li><a\nhref=\"https://github.com/vercel/next.js/commit/be4aafd4b744fbf6500b311b74c84243f70a3059\"><code>be4aafd</code></a>\nv15.4.6</li>\n<li><a\nhref=\"https://github.com/vercel/next.js/commit/91e5b6b84f56c096dc2bb647eb384388e06489fd\"><code>91e5b6b</code></a>\nBackport &quot;fix: add <code>?dpl</code> to fonts in\n<code>/_next/static/media</code> (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82384\">#82384</a>)&quot;\n(<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82421\">#82421</a>)</li>\n<li><a\nhref=\"https://github.com/vercel/next.js/commit/f1629d939597cc46ccbe44fe66bda91eac31e219\"><code>f1629d9</code></a>\nBackport &quot;[Pages] fix: <code>_error</code> page's\n<code>req.url</code> can be overwritten t… (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82377\">#82377</a>)</li>\n<li><a\nhref=\"https://github.com/vercel/next.js/commit/b9aab5dbe926c256d439bfecb226693dcd1a1be4\"><code>b9aab5d</code></a>\nv15.4.5</li>\n<li><a\nhref=\"https://github.com/vercel/next.js/commit/a8c93c49dd2d42a2ced8a17bc53b3fea11d25c96\"><code>a8c93c4</code></a>\nDisable test new tests jobs</li>\n<li><a\nhref=\"https://github.com/vercel/next.js/commit/ed2a6c754831406f8cd724eb52a562bb124c1504\"><code>ed2a6c7</code></a>\n[backport]: fix(next/image): improve and simplify detect-content-type\n(<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82118\">#82118</a>...</li>\n<li><a\nhref=\"https://github.com/vercel/next.js/commit/f00fcc9011e7d0ef027021fc8424f51b8ac97880\"><code>f00fcc9</code></a>\n[backport]: fix(next/image): fix image-optimizer.ts headers (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82114\">#82114</a>)\n(<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82175\">#82175</a>)</li>\n<li><a\nhref=\"https://github.com/vercel/next.js/commit/55a7568e9d12459f8c5f5ae120fe2e228af284bb\"><code>55a7568</code></a>\nBackport: Turbopack: update mimalloc (<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/81993\">#81993</a>)\n(<a\nhref=\"https://redirect.github.com/vercel/next.js/issues/82166\">#82166</a>)</li>\n<li>Additional commits viewable in <a\nhref=\"https://github.com/vercel/next.js/compare/v15.3.3...v15.4.7\">compare\nview</a></li>\n</ul>\n</details>\n<br />\n\n\n[![Dependabot compatibility\nscore](https://dependabot-badges.githubapp.com/badges/compatibility_score?dependency-name=next&package-manager=npm_and_yarn&previous-version=15.3.3&new-version=15.4.7)](https://docs.github.com/en/github/managing-security-vulnerabilities/about-dependabot-security-updates#about-compatibility-scores)\n\nDependabot will resolve any conflicts with this PR as long as you don't\nalter it yourself. You can also trigger a rebase manually by commenting\n`@dependabot rebase`.\n\n[//]: # (dependabot-automerge-start)\n[//]: # (dependabot-automerge-end)\n\n---\n\n<details>\n<summary>Dependabot commands and options</summary>\n<br />\n\nYou can trigger Dependabot actions by commenting on this PR:\n- `@dependabot rebase` will rebase this PR\n- `@dependabot recreate` will recreate this PR, overwriting any edits\nthat have been made to it\n- `@dependabot merge` will merge this PR after your CI passes on it\n- `@dependabot squash and merge` will squash and merge this PR after\nyour CI passes on it\n- `@dependabot cancel merge` will cancel a previously requested merge\nand block automerging\n- `@dependabot reopen` will reopen this PR if it is closed\n- `@dependabot close` will close this PR and stop Dependabot recreating\nit. You can achieve the same result by closing it manually\n- `@dependabot show <dependency name> ignore conditions` will show all\nof the ignore conditions of the specified dependency\n- `@dependabot ignore this major version` will close this PR and stop\nDependabot creating any more for this major version (unless you reopen\nthe PR or upgrade to it yourself)\n- `@dependabot ignore this minor version` will close this PR and stop\nDependabot creating any more for this minor version (unless you reopen\nthe PR or upgrade to it yourself)\n- `@dependabot ignore this dependency` will close this PR and stop\nDependabot creating any more for this dependency (unless you reopen the\nPR or upgrade to it yourself)\nYou can disable automated security fix PRs for this repo from the\n[Security Alerts\npage](https://github.com/AzureAD/microsoft-authentication-library-for-js/network/alerts).\n\n</details>\n\nSigned-off-by: dependabot[bot] <support@github.com>\nCo-authored-by: dependabot[bot] <49699333+dependabot[bot]@users.noreply.github.com>",
          "timestamp": "2025-09-03T10:19:39-07:00",
          "tree_id": "403066bcbc3fcadd0750237e8be85ed81c02a03b",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/5b86b0bdacc0b89a714b5a977dd19c63a80f4248"
        },
        "date": 1756939284414,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 239636,
            "range": "±1.19%",
            "unit": "ops/sec",
            "extra": "223 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 250421,
            "range": "±0.83%",
            "unit": "ops/sec",
            "extra": "224 samples"
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
          "id": "39fa40f3f6dfd3ef91493e04c7f80f636a765e97",
          "message": "Release PR: official (#8028)\n\nThis PR contains the changelogs and version bumps for the MSAL.js 3P\nreleases.\n\nCo-authored-by: MSAL.js Release Automation <msaljsbuilds@microsoft.com>\nCo-authored-by: Hector Morales <hemoral@microsoft.com>",
          "timestamp": "2025-09-04T10:03:13-07:00",
          "tree_id": "4794df07ce9dc952c292ae893d06a45b795bef8c",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/39fa40f3f6dfd3ef91493e04c7f80f636a765e97"
        },
        "date": 1757020485127,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 249245,
            "range": "±0.85%",
            "unit": "ops/sec",
            "extra": "237 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 249471,
            "range": "±0.72%",
            "unit": "ops/sec",
            "extra": "235 samples"
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
          "id": "48114c12a879c48213bb34cb45c87c316cf3bc88",
          "message": "Export client capabilities helper (#8033)\n\nExports client capabilities helper to be used by the pairwise broker\nwhen crafting embedded request",
          "timestamp": "2025-09-08T09:46:00-07:00",
          "tree_id": "58296faacc173b47d4f5d5efefde5a9e42fc5bbe",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/48114c12a879c48213bb34cb45c87c316cf3bc88"
        },
        "date": 1757367346763,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 249208,
            "range": "±0.84%",
            "unit": "ops/sec",
            "extra": "234 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 253494,
            "range": "±0.71%",
            "unit": "ops/sec",
            "extra": "235 samples"
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
          "id": "bd9cfdcecfd1c3903f3e04ec70466b8794de2cd6",
          "message": "Bump actions/labeler from 4 to 6 (#8036)\n\nBumps [actions/labeler](https://github.com/actions/labeler) from 4 to 6.\n<details>\n<summary>Release notes</summary>\n<p><em>Sourced from <a\nhref=\"https://github.com/actions/labeler/releases\">actions/labeler's\nreleases</a>.</em></p>\n<blockquote>\n<h2>v6.0.0</h2>\n<h2>What's Changed</h2>\n<ul>\n<li>Add workflow file for publishing releases to immutable action\npackage by <a\nhref=\"https://github.com/jcambass\"><code>@​jcambass</code></a> in <a\nhref=\"https://redirect.github.com/actions/labeler/pull/802\">actions/labeler#802</a></li>\n</ul>\n<h3>Breaking Changes</h3>\n<ul>\n<li>Upgrade Node.js version to 24 in action and dependencies <a\nhref=\"https://github.com/salmanmkc\"><code>@​salmanmkc</code></a> in <a\nhref=\"https://redirect.github.com/actions/labeler/pull/891\">actions/labeler#891</a>\nMake sure your runner is on version v2.327.1 or later to ensure\ncompatibility with this release. <a\nhref=\"https://github.com/actions/runner/releases/tag/v2.327.1\">Release\nNotes</a></li>\n</ul>\n<h3>Dependency Upgrades</h3>\n<ul>\n<li>Upgrade eslint-config-prettier from 9.0.0 to 9.1.0 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/711\">actions/labeler#711</a></li>\n<li>Upgrade eslint from 8.52.0 to 8.55.0 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/720\">actions/labeler#720</a></li>\n<li>Upgrade <code>@​types/jest</code> from 29.5.6 to 29.5.11 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/719\">actions/labeler#719</a></li>\n<li>Upgrade <code>@​types/js-yaml</code> from 4.0.8 to 4.0.9 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/718\">actions/labeler#718</a></li>\n<li>Upgrade <code>@​typescript-eslint/parser</code> from 6.9.0 to 6.14.0\nby <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/717\">actions/labeler#717</a></li>\n<li>Upgrade prettier from 3.0.3 to 3.1.1 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/726\">actions/labeler#726</a></li>\n<li>Upgrade eslint from 8.55.0 to 8.56.0 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/725\">actions/labeler#725</a></li>\n<li>Upgrade <code>@​typescript-eslint/parser</code> from 6.14.0 to\n6.19.0 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/745\">actions/labeler#745</a></li>\n<li>Upgrade eslint-plugin-jest from 27.4.3 to 27.6.3 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/744\">actions/labeler#744</a></li>\n<li>Upgrade <code>@​typescript-eslint/eslint-plugin</code> from 6.9.0 to\n6.20.0 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/750\">actions/labeler#750</a></li>\n<li>Upgrade prettier from 3.1.1 to 3.2.5 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/752\">actions/labeler#752</a></li>\n<li>Upgrade undici from 5.26.5 to 5.28.3 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/757\">actions/labeler#757</a></li>\n<li>Upgrade braces from 3.0.2 to 3.0.3 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/789\">actions/labeler#789</a></li>\n<li>Upgrade minimatch from 9.0.3 to 10.0.1 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/805\">actions/labeler#805</a></li>\n<li>Upgrade <code>@​actions/core</code> from 1.10.1 to 1.11.1 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/811\">actions/labeler#811</a></li>\n<li>Upgrade typescript from 5.4.3 to 5.7.2 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/819\">actions/labeler#819</a></li>\n<li>Upgrade <code>@​typescript-eslint/parser</code> from 7.3.1 to 8.17.0\nby <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/824\">actions/labeler#824</a></li>\n<li>Upgrade prettier from 3.2.5 to 3.4.2 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/825\">actions/labeler#825</a></li>\n<li>Upgrade <code>@​types/jest</code> from 29.5.12 to 29.5.14 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/827\">actions/labeler#827</a></li>\n<li>Upgrade eslint-plugin-jest from 27.9.0 to 28.9.0 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/832\">actions/labeler#832</a></li>\n<li>Upgrade ts-jest from 29.1.2 to 29.2.5 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/831\">actions/labeler#831</a></li>\n<li>Upgrade <code>@​vercel/ncc</code> from 0.38.1 to 0.38.3 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/830\">actions/labeler#830</a></li>\n<li>Upgrade typescript from 5.7.2 to 5.7.3 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/835\">actions/labeler#835</a></li>\n<li>Upgrade eslint-plugin-jest from 28.9.0 to 28.11.0 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/839\">actions/labeler#839</a></li>\n<li>Upgrade undici from 5.28.4 to 5.28.5 by <a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/842\">actions/labeler#842</a></li>\n<li>Upgrade <code>@​octokit/request-error</code> from 5.0.1 to 5.1.1 by\n<a\nhref=\"https://github.com/dependabot\"><code>@​dependabot</code></a>[bot]\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/846\">actions/labeler#846</a></li>\n</ul>\n<h3>Documentation changes</h3>\n<ul>\n<li>Add note regarding <code>pull_request_target</code> to README.md by\n<a href=\"https://github.com/silverwind\"><code>@​silverwind</code></a> in\n<a\nhref=\"https://redirect.github.com/actions/labeler/pull/669\">actions/labeler#669</a></li>\n<li>Update readme with additional examples and important note about\n<code>pull_request_target</code> event by <a\nhref=\"https://github.com/IvanZosimov\"><code>@​IvanZosimov</code></a> in\n<a\nhref=\"https://redirect.github.com/actions/labeler/pull/721\">actions/labeler#721</a></li>\n<li>Document update - permission section by <a\nhref=\"https://github.com/harithavattikuti\"><code>@​harithavattikuti</code></a>\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/840\">actions/labeler#840</a></li>\n<li>Improvement in documentation for pull_request_target event usage in\nREADME by <a\nhref=\"https://github.com/suyashgaonkar\"><code>@​suyashgaonkar</code></a>\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/871\">actions/labeler#871</a></li>\n<li>Fix broken links in documentation by <a\nhref=\"https://github.com/suyashgaonkar\"><code>@​suyashgaonkar</code></a>\nin <a\nhref=\"https://redirect.github.com/actions/labeler/pull/822\">actions/labeler#822</a></li>\n</ul>\n<h2>New Contributors</h2>\n<ul>\n<li><a\nhref=\"https://github.com/silverwind\"><code>@​silverwind</code></a> made\ntheir first contribution in <a\nhref=\"https://redirect.github.com/actions/labeler/pull/669\">actions/labeler#669</a></li>\n<li><a href=\"https://github.com/Jcambass\"><code>@​Jcambass</code></a>\nmade their first contribution in <a\nhref=\"https://redirect.github.com/actions/labeler/pull/802\">actions/labeler#802</a></li>\n<li><a\nhref=\"https://github.com/suyashgaonkar\"><code>@​suyashgaonkar</code></a>\nmade their first contribution in <a\nhref=\"https://redirect.github.com/actions/labeler/pull/822\">actions/labeler#822</a></li>\n<li><a\nhref=\"https://github.com/HarithaVattikuti\"><code>@​HarithaVattikuti</code></a>\nmade their first contribution in <a\nhref=\"https://redirect.github.com/actions/labeler/pull/840\">actions/labeler#840</a></li>\n<li><a href=\"https://github.com/salmanmkc\"><code>@​salmanmkc</code></a>\nmade their first contribution in <a\nhref=\"https://redirect.github.com/actions/labeler/pull/891\">actions/labeler#891</a></li>\n</ul>\n<!-- raw HTML omitted -->\n</blockquote>\n<p>... (truncated)</p>\n</details>\n<details>\n<summary>Commits</summary>\n<ul>\n<li><a\nhref=\"https://github.com/actions/labeler/commit/634933edcd8ababfe52f92936142cc22ac488b1b\"><code>634933e</code></a>\npublish-action upgrade to 0.4.0 from 0.2.2 (<a\nhref=\"https://redirect.github.com/actions/labeler/issues/901\">#901</a>)</li>\n<li><a\nhref=\"https://github.com/actions/labeler/commit/f1a63e87db0c6baf19c5713083f8d00d789ca184\"><code>f1a63e8</code></a>\nUpdate Node.js version to 24 in action and dependencies (<a\nhref=\"https://redirect.github.com/actions/labeler/issues/891\">#891</a>)</li>\n<li><a\nhref=\"https://github.com/actions/labeler/commit/b0a1180683c9f17424de4d71c044bea4c7b9bc7c\"><code>b0a1180</code></a>\nBump <code>@​octokit/request-error</code> from 5.0.1 to 5.1.1 (<a\nhref=\"https://redirect.github.com/actions/labeler/issues/846\">#846</a>)</li>\n<li><a\nhref=\"https://github.com/actions/labeler/commit/110d44140c9195b853f2f24044bbfed8f4968efb\"><code>110d441</code></a>\nUpdate README.md (<a\nhref=\"https://redirect.github.com/actions/labeler/issues/871\">#871</a>)</li>\n<li><a\nhref=\"https://github.com/actions/labeler/commit/bee50fefe18762fad67754b2f3bfff2c8082ebb8\"><code>bee50fe</code></a>\nBump undici from 5.28.4 to 5.28.5 (<a\nhref=\"https://redirect.github.com/actions/labeler/issues/842\">#842</a>)</li>\n<li><a\nhref=\"https://github.com/actions/labeler/commit/6463cdb00ee92c05bec55dffc4e1fce250301945\"><code>6463cdb</code></a>\nBump eslint-plugin-jest from 28.9.0 to 28.11.0 (<a\nhref=\"https://redirect.github.com/actions/labeler/issues/839\">#839</a>)</li>\n<li><a\nhref=\"https://github.com/actions/labeler/commit/c209686724ee12fcc5e6294d1d569b91f86fa691\"><code>c209686</code></a>\nBump typescript from 5.7.2 to 5.7.3 (<a\nhref=\"https://redirect.github.com/actions/labeler/issues/835\">#835</a>)</li>\n<li><a\nhref=\"https://github.com/actions/labeler/commit/5184940b544b0096088a7b42d1b8a551003d9eb1\"><code>5184940</code></a>\nBump <code>@​vercel/ncc</code> from 0.38.1 to 0.38.3 (<a\nhref=\"https://redirect.github.com/actions/labeler/issues/830\">#830</a>)</li>\n<li><a\nhref=\"https://github.com/actions/labeler/commit/3629d5568b59204f18786372f6d740d649719488\"><code>3629d55</code></a>\nDocument update - permission section (<a\nhref=\"https://redirect.github.com/actions/labeler/issues/840\">#840</a>)</li>\n<li><a\nhref=\"https://github.com/actions/labeler/commit/d24f7f3731b2a06433c0bccc364d560c5329c48f\"><code>d24f7f3</code></a>\nBump ts-jest from 29.1.2 to 29.2.5 (<a\nhref=\"https://redirect.github.com/actions/labeler/issues/831\">#831</a>)</li>\n<li>Additional commits viewable in <a\nhref=\"https://github.com/actions/labeler/compare/v4...v6\">compare\nview</a></li>\n</ul>\n</details>\n<br />\n\n\n[![Dependabot compatibility\nscore](https://dependabot-badges.githubapp.com/badges/compatibility_score?dependency-name=actions/labeler&package-manager=github_actions&previous-version=4&new-version=6)](https://docs.github.com/en/github/managing-security-vulnerabilities/about-dependabot-security-updates#about-compatibility-scores)\n\nDependabot will resolve any conflicts with this PR as long as you don't\nalter it yourself. You can also trigger a rebase manually by commenting\n`@dependabot rebase`.\n\n[//]: # (dependabot-automerge-start)\n[//]: # (dependabot-automerge-end)\n\n---\n\n<details>\n<summary>Dependabot commands and options</summary>\n<br />\n\nYou can trigger Dependabot actions by commenting on this PR:\n- `@dependabot rebase` will rebase this PR\n- `@dependabot recreate` will recreate this PR, overwriting any edits\nthat have been made to it\n- `@dependabot merge` will merge this PR after your CI passes on it\n- `@dependabot squash and merge` will squash and merge this PR after\nyour CI passes on it\n- `@dependabot cancel merge` will cancel a previously requested merge\nand block automerging\n- `@dependabot reopen` will reopen this PR if it is closed\n- `@dependabot close` will close this PR and stop Dependabot recreating\nit. You can achieve the same result by closing it manually\n- `@dependabot show <dependency name> ignore conditions` will show all\nof the ignore conditions of the specified dependency\n- `@dependabot ignore this major version` will close this PR and stop\nDependabot creating any more for this major version (unless you reopen\nthe PR or upgrade to it yourself)\n- `@dependabot ignore this minor version` will close this PR and stop\nDependabot creating any more for this minor version (unless you reopen\nthe PR or upgrade to it yourself)\n- `@dependabot ignore this dependency` will close this PR and stop\nDependabot creating any more for this dependency (unless you reopen the\nPR or upgrade to it yourself)\n\n\n</details>\n\nSigned-off-by: dependabot[bot] <support@github.com>\nCo-authored-by: dependabot[bot] <49699333+dependabot[bot]@users.noreply.github.com>\nCo-authored-by: Thomas Norling <thomas.norling@microsoft.com>",
          "timestamp": "2025-09-09T10:40:30-07:00",
          "tree_id": "9f70d4d2749e534374b5e9dd507538c370bdbd56",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/bd9cfdcecfd1c3903f3e04ec70466b8794de2cd6"
        },
        "date": 1757455844525,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 252825,
            "range": "±1.11%",
            "unit": "ops/sec",
            "extra": "236 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 252535,
            "range": "±0.77%",
            "unit": "ops/sec",
            "extra": "236 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "msaljsbuilds@microsoft.com",
            "name": "MSAL.js Release Automation"
          },
          "committer": {
            "email": "kshabelko@microsoft.com",
            "name": "Konstantin",
            "username": "konstantin-msft"
          },
          "distinct": true,
          "id": "e341099149645debcdbc0632cd3e17b563eb351a",
          "message": "Bump package versions",
          "timestamp": "2025-09-09T17:54:59-04:00",
          "tree_id": "947a5d9daea53ece9258489638993840df202eac",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/e341099149645debcdbc0632cd3e17b563eb351a"
        },
        "date": 1757483996945,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 255065,
            "range": "±0.79%",
            "unit": "ops/sec",
            "extra": "235 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 255519,
            "range": "±0.58%",
            "unit": "ops/sec",
            "extra": "236 samples"
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
          "id": "8117dc6664bf77652457ba1a58b29df6402890f6",
          "message": "Change files",
          "timestamp": "2025-09-10T10:48:59-04:00",
          "tree_id": "22fce366f103b87d5c4359b418939c3d906fbf08",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/8117dc6664bf77652457ba1a58b29df6402890f6"
        },
        "date": 1757525569812,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 253003,
            "range": "±0.94%",
            "unit": "ops/sec",
            "extra": "215 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 249444,
            "range": "±0.72%",
            "unit": "ops/sec",
            "extra": "233 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "45603194+shenj@users.noreply.github.com",
            "name": "Jian Shen",
            "username": "shenj"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "b66a707fdfa0c46c3f367b46e8c756985bfea3c1",
          "message": "Support JIT during Sign-In and Add AI instructions files (#8007)\n\nThis PR merges branch custom-auth/jiashen/jit into dev.\nSummary:\n- Support JIT experience during Sign-In in msal-browser/custom_auth\nflows\n- Add .github Copilot instructions and custom_auth instructions files",
          "timestamp": "2025-09-11T09:37:28+01:00",
          "tree_id": "602bd3dea7a29a5fe2a2d8ed3049a1654e10bd50",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/b66a707fdfa0c46c3f367b46e8c756985bfea3c1"
        },
        "date": 1757582079252,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 248279,
            "range": "±1.11%",
            "unit": "ops/sec",
            "extra": "223 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 246880,
            "range": "±0.83%",
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
          "id": "678d35f1842fe1b1f160d965fbffe1ae234ec2c4",
          "message": "- Fix formatting",
          "timestamp": "2025-09-11T16:57:22-04:00",
          "tree_id": "f5a8b3936ef4fe5e94db34e5665b2ecaf5f0588a",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/678d35f1842fe1b1f160d965fbffe1ae234ec2c4"
        },
        "date": 1757625293056,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 240365,
            "range": "±0.91%",
            "unit": "ops/sec",
            "extra": "233 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 256818,
            "range": "±0.71%",
            "unit": "ops/sec",
            "extra": "239 samples"
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
          "id": "f65f2dc0dd1315acfd008b309bb5f0a1e00dd7a7",
          "message": "Bump element-plus from 2.9.3 to 2.11.0 (#8043)\n\nBumps [element-plus](https://github.com/element-plus/element-plus) from\n2.9.3 to 2.11.0.\n<details>\n<summary>Release notes</summary>\n<p><em>Sourced from <a\nhref=\"https://github.com/element-plus/element-plus/releases\">element-plus's\nreleases</a>.</em></p>\n<blockquote>\n<h3>2.11.0</h3>\n</blockquote>\n<blockquote>\n</blockquote>\n<blockquote>\n<p><em>2025-08-22</em></p>\n</blockquote>\n<blockquote>\n</blockquote>\n<blockquote>\n<h4>Features</h4>\n</blockquote>\n<blockquote>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [tour] support keyboard switching steps (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21766\">#21766</a>\nby <a\nhref=\"https://github.com/thinkasany\"><code>@​thinkasany</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [select] add support for option groups in\n<code>options</code> (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21692\">#21692</a>\nby <a href=\"https://github.com/cszhjh\"><code>@​cszhjh</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components add custom clear icon (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21699\">#21699</a>\nby <a\nhref=\"https://github.com/a869246700\"><code>@​a869246700</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components add isDragging to useDraggable for dialog/message-box (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21703\">#21703</a>\nby <a\nhref=\"https://github.com/zhixiaotong\"><code>@​zhixiaotong</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [splitter] add <code>lazy</code> prop (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21730\">#21730</a>\nby <a href=\"https://github.com/cszhjh\"><code>@​cszhjh</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [drawer] add <code>resizable</code> prop (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21608\">#21608</a>\nby <a href=\"https://github.com/cszhjh\"><code>@​cszhjh</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components unify clear btn display on focus (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21732\">#21732</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [date-picker-panel] externalize <code>date-picker</code>\ncore component (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21331\">#21331</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [config-provider/button] add text prop (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21806\">#21806</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [message] add <code>placement</code> option &amp; method\n(<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21747\">#21747</a>\nby <a\nhref=\"https://github.com/zhixiaotong\"><code>@​zhixiaotong</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [cascader/select/date-picker] add popper-style (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21807\">#21807</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [input-tag] add <code>collapse-*</code> prop (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21477\">#21477</a>\nby <a\nhref=\"https://github.com/eighty-cents\"><code>@​eighty-cents</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [color-picker-panel] externalize\n<code>color-picker</code> core component (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21668\">#21668</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n</blockquote>\n<blockquote>\n<h4>Bug fixes</h4>\n</blockquote>\n<blockquote>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [carousel] special case <code>activeIndex</code> is\ninaccurate (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21736\">#21736</a>\nby <a href=\"https://github.com/jevin98\"><code>@​jevin98</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [splitter] <code>collapsible</code> does not take effect\ndynamically (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21742\">#21742</a>\nby <a\nhref=\"https://github.com/thinkasany\"><code>@​thinkasany</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [cascader-panel] ensure set null after clear (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21652\">#21652</a>\nby <a\nhref=\"https://github.com/snowbitx\"><code>@​snowbitx</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [menu] skip comment nodes in ellipsis (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21751\">#21751</a>\nby <a href=\"https://github.com/cszhjh\"><code>@​cszhjh</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [cascader] issue when lazy loading returns leaf nodes (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21414\">#21414</a>\nby <a\nhref=\"https://github.com/keeplearning66\"><code>@​keeplearning66</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [splitter] collapse works with two-way size and min &amp;\nmax (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21738\">#21738</a>\nby <a href=\"https://github.com/cszhjh\"><code>@​cszhjh</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [segmented] wrap options item slot to any (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21794\">#21794</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [steps] dynamic status step exception (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21753\">#21753</a>\nby <a href=\"https://github.com/dhj-l\"><code>@​dhj-l</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [cascader] correctly handle focus and blur events (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21510\">#21510</a>\nby <a href=\"https://github.com/cszhjh\"><code>@​cszhjh</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [input] add <code>work-break</code> to\n<code>CONTEXT_STYLE</code> (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21809\">#21809</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [slider] correctly reset sliderSize (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21812\">#21812</a>\nby <a href=\"https://github.com/cszhjh\"><code>@​cszhjh</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [radio-group] prop-value cause ivalidate loop (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21814\">#21814</a>\nby <a\nhref=\"https://github.com/cactuser-Lu\"><code>@​cactuser-Lu</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [table] make root <code>show-overflow-tooltip</code>\ndynamic (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21823\">#21823</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [table] show-overflow-tooltip dynamic switch (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21811\">#21811</a>\nby <a href=\"https://github.com/dhj-l\"><code>@​dhj-l</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Build path compatible with Windows platform (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21423\">#21423</a>\nby <a\nhref=\"https://github.com/Monday1997\"><code>@​Monday1997</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [cascader] treat <code>tagTooltipRef</code> as array (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21826\">#21826</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [date-picker] add missing data slot (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/19839\">#19839</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [select &amp; select-v2] avoid line break when input is\nempty (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21844\">#21844</a>\nby <a\nhref=\"https://github.com/micaiguai\"><code>@​micaiguai</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [color-picker-panel] sync bg-color with date-picker-panel\n(<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21859\">#21859</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n</blockquote>\n<blockquote>\n<h4>Refactors</h4>\n</blockquote>\n<blockquote>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components export <code>TreeV2Instance</code> &amp; optimize type\nusage (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21718\">#21718</a>\nby <a\nhref=\"https://github.com/zhixiaotong\"><code>@​zhixiaotong</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<ul>\n<li>Components [cascader-panel] unify click node (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21796\">#21796</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n</ul>\n</blockquote>\n<blockquote>\n<h3>2.10.7</h3>\n<p><em>2025-08-12</em></p>\n<!-- raw HTML omitted -->\n</blockquote>\n<p>... (truncated)</p>\n</details>\n<details>\n<summary>Changelog</summary>\n<p><em>Sourced from <a\nhref=\"https://github.com/element-plus/element-plus/blob/dev/CHANGELOG.en-US.md\">element-plus's\nchangelog</a>.</em></p>\n<blockquote>\n<h3>2.11.0</h3>\n<p><em>2025-08-22</em></p>\n<h4>Features</h4>\n<ul>\n<li>Components [tour] support keyboard switching steps (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21766\">#21766</a>\nby <a\nhref=\"https://github.com/thinkasany\"><code>@​thinkasany</code></a>)</li>\n<li>Components [select] add support for option groups in\n<code>options</code> (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21692\">#21692</a>\nby <a href=\"https://github.com/cszhjh\"><code>@​cszhjh</code></a>)</li>\n<li>Components add custom clear icon (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21699\">#21699</a>\nby <a\nhref=\"https://github.com/a869246700\"><code>@​a869246700</code></a>)</li>\n<li>Components add isDragging to useDraggable for dialog/message-box (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21703\">#21703</a>\nby <a\nhref=\"https://github.com/zhixiaotong\"><code>@​zhixiaotong</code></a>)</li>\n<li>Components [splitter] add <code>lazy</code> prop (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21730\">#21730</a>\nby <a href=\"https://github.com/cszhjh\"><code>@​cszhjh</code></a>)</li>\n<li>Components [drawer] add <code>resizable</code> prop (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21608\">#21608</a>\nby <a href=\"https://github.com/cszhjh\"><code>@​cszhjh</code></a>)</li>\n<li>Components unify clear btn display on focus (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21732\">#21732</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n<li>Components [date-picker-panel] externalize <code>date-picker</code>\ncore component (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21331\">#21331</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n<li>Components [config-provider/button] add text prop (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21806\">#21806</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n<li>Components [message] add <code>placement</code> option &amp; method\n(<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21747\">#21747</a>\nby <a\nhref=\"https://github.com/zhixiaotong\"><code>@​zhixiaotong</code></a>)</li>\n<li>Components [cascader/select/date-picker] add popper-style (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21807\">#21807</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n<li>Components [input-tag] add <code>collapse-*</code> prop (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21477\">#21477</a>\nby <a\nhref=\"https://github.com/eighty-cents\"><code>@​eighty-cents</code></a>)</li>\n<li>Components [color-picker-panel] externalize\n<code>color-picker</code> core component (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21668\">#21668</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n</ul>\n<h4>Bug fixes</h4>\n<ul>\n<li>Components [carousel] special case <code>activeIndex</code> is\ninaccurate (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21736\">#21736</a>\nby <a href=\"https://github.com/jevin98\"><code>@​jevin98</code></a>)</li>\n<li>Components [splitter] <code>collapsible</code> does not take effect\ndynamically (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21742\">#21742</a>\nby <a\nhref=\"https://github.com/thinkasany\"><code>@​thinkasany</code></a>)</li>\n<li>Components [cascader-panel] ensure set null after clear (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21652\">#21652</a>\nby <a\nhref=\"https://github.com/snowbitx\"><code>@​snowbitx</code></a>)</li>\n<li>Components [menu] skip comment nodes in ellipsis (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21751\">#21751</a>\nby <a href=\"https://github.com/cszhjh\"><code>@​cszhjh</code></a>)</li>\n<li>Components [cascader] issue when lazy loading returns leaf nodes (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21414\">#21414</a>\nby <a\nhref=\"https://github.com/keeplearning66\"><code>@​keeplearning66</code></a>)</li>\n<li>Components [splitter] collapse works with two-way size and min &amp;\nmax (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21738\">#21738</a>\nby <a href=\"https://github.com/cszhjh\"><code>@​cszhjh</code></a>)</li>\n<li>Components [segmented] wrap options item slot to any (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21794\">#21794</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n<li>Components [steps] dynamic status step exception (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21753\">#21753</a>\nby <a href=\"https://github.com/dhj-l\"><code>@​dhj-l</code></a>)</li>\n<li>Components [cascader] correctly handle focus and blur events (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21510\">#21510</a>\nby <a href=\"https://github.com/cszhjh\"><code>@​cszhjh</code></a>)</li>\n<li>Components [input] add <code>work-break</code> to\n<code>CONTEXT_STYLE</code> (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21809\">#21809</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n<li>Components [slider] correctly reset sliderSize (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21812\">#21812</a>\nby <a href=\"https://github.com/cszhjh\"><code>@​cszhjh</code></a>)</li>\n<li>Components [radio-group] prop-value cause ivalidate loop (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21814\">#21814</a>\nby <a\nhref=\"https://github.com/cactuser-Lu\"><code>@​cactuser-Lu</code></a>)</li>\n<li>Components [table] make root <code>show-overflow-tooltip</code>\ndynamic (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21823\">#21823</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n<li>Components [table] show-overflow-tooltip dynamic switch (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21811\">#21811</a>\nby <a href=\"https://github.com/dhj-l\"><code>@​dhj-l</code></a>)</li>\n<li>Build path compatible with Windows platform (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21423\">#21423</a>\nby <a\nhref=\"https://github.com/Monday1997\"><code>@​Monday1997</code></a>)</li>\n<li>Components [cascader] treat <code>tagTooltipRef</code> as array (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21826\">#21826</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n<li>Components [date-picker] add missing data slot (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/19839\">#19839</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n<li>Components [select &amp; select-v2] avoid line break when input is\nempty (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21844\">#21844</a>\nby <a\nhref=\"https://github.com/micaiguai\"><code>@​micaiguai</code></a>)</li>\n<li>Components [color-picker-panel] sync bg-color with date-picker-panel\n(<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21859\">#21859</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n</ul>\n<h4>Refactors</h4>\n<ul>\n<li>Components export <code>TreeV2Instance</code> &amp; optimize type\nusage (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21718\">#21718</a>\nby <a\nhref=\"https://github.com/zhixiaotong\"><code>@​zhixiaotong</code></a>)</li>\n<li>Components [cascader-panel] unify click node (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21796\">#21796</a>\nby <a href=\"https://github.com/Dsaquel\"><code>@​Dsaquel</code></a>)</li>\n</ul>\n<h3>2.10.7</h3>\n<p><em>2025-08-12</em></p>\n<!-- raw HTML omitted -->\n</blockquote>\n<p>... (truncated)</p>\n</details>\n<details>\n<summary>Commits</summary>\n<ul>\n<li><a\nhref=\"https://github.com/element-plus/element-plus/commit/d8a3eae1a164c0ac67fe64d474b2338d70de7452\"><code>d8a3eae</code></a>\nMerge pull request <a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21863\">#21863</a>\nfrom element-plus/dev</li>\n<li><a\nhref=\"https://github.com/element-plus/element-plus/commit/9181afa895c6943a78a3537581e88fb01b99d710\"><code>9181afa</code></a>\nchore: Update changelog 2.11.0 (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21864\">#21864</a>)</li>\n<li><a\nhref=\"https://github.com/element-plus/element-plus/commit/289b59e546d87fa888b1bb2c51326ddbcba173a5\"><code>289b59e</code></a>\nfix(components): [color-picker-panel] sync bg-color with\ndate-picker-panel (#...</li>\n<li><a\nhref=\"https://github.com/element-plus/element-plus/commit/2ccb268e823793cf88becc3f20fb17d42058a0c7\"><code>2ccb268</code></a>\nfix(components): [select &amp; select-v2] avoid line break when input is\nempty (#...</li>\n<li><a\nhref=\"https://github.com/element-plus/element-plus/commit/2144a78880fd5557b075e37d10295921db89fc3f\"><code>2144a78</code></a>\nfeat(components): [color-picker-panel] externalize\n<code>color-picker</code> core compon...</li>\n<li><a\nhref=\"https://github.com/element-plus/element-plus/commit/2e19e59e67cc55f6b2febff156330f86d9f75415\"><code>2e19e59</code></a>\nfix(components): [date-picker] add missing data slot (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/19839\">#19839</a>)</li>\n<li><a\nhref=\"https://github.com/element-plus/element-plus/commit/5647b3e392f8c0ccaf273756b9e3a8ef1efb2a28\"><code>5647b3e</code></a>\nfeat(components): [input-tag] add <code>collapse-*</code> prop (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21477\">#21477</a>)</li>\n<li><a\nhref=\"https://github.com/element-plus/element-plus/commit/a985f511a0ca43d2e8972c4efa9ba69022491031\"><code>a985f51</code></a>\nchore: date-picker-panel tweaks (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21827\">#21827</a>)</li>\n<li><a\nhref=\"https://github.com/element-plus/element-plus/commit/96c5d8ae8cf5b7a149a682216ecaf0634df97aa0\"><code>96c5d8a</code></a>\nfix(components): [cascader] treat <code>tagTooltipRef</code> as array\n(<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21826\">#21826</a>)</li>\n<li><a\nhref=\"https://github.com/element-plus/element-plus/commit/7644c1fe4a2e7ff7fb7fe02680c0d84583cac3cc\"><code>7644c1f</code></a>\nfeat(components): [cascader/select/date-picker] add popper-style (<a\nhref=\"https://redirect.github.com/element-plus/element-plus/issues/21807\">#21807</a>)</li>\n<li>Additional commits viewable in <a\nhref=\"https://github.com/element-plus/element-plus/compare/2.9.3...2.11.0\">compare\nview</a></li>\n</ul>\n</details>\n<br />\n\n\n[![Dependabot compatibility\nscore](https://dependabot-badges.githubapp.com/badges/compatibility_score?dependency-name=element-plus&package-manager=npm_and_yarn&previous-version=2.9.3&new-version=2.11.0)](https://docs.github.com/en/github/managing-security-vulnerabilities/about-dependabot-security-updates#about-compatibility-scores)\n\nDependabot will resolve any conflicts with this PR as long as you don't\nalter it yourself. You can also trigger a rebase manually by commenting\n`@dependabot rebase`.\n\n[//]: # (dependabot-automerge-start)\n[//]: # (dependabot-automerge-end)\n\n---\n\n<details>\n<summary>Dependabot commands and options</summary>\n<br />\n\nYou can trigger Dependabot actions by commenting on this PR:\n- `@dependabot rebase` will rebase this PR\n- `@dependabot recreate` will recreate this PR, overwriting any edits\nthat have been made to it\n- `@dependabot merge` will merge this PR after your CI passes on it\n- `@dependabot squash and merge` will squash and merge this PR after\nyour CI passes on it\n- `@dependabot cancel merge` will cancel a previously requested merge\nand block automerging\n- `@dependabot reopen` will reopen this PR if it is closed\n- `@dependabot close` will close this PR and stop Dependabot recreating\nit. You can achieve the same result by closing it manually\n- `@dependabot show <dependency name> ignore conditions` will show all\nof the ignore conditions of the specified dependency\n- `@dependabot ignore this major version` will close this PR and stop\nDependabot creating any more for this major version (unless you reopen\nthe PR or upgrade to it yourself)\n- `@dependabot ignore this minor version` will close this PR and stop\nDependabot creating any more for this minor version (unless you reopen\nthe PR or upgrade to it yourself)\n- `@dependabot ignore this dependency` will close this PR and stop\nDependabot creating any more for this dependency (unless you reopen the\nPR or upgrade to it yourself)\nYou can disable automated security fix PRs for this repo from the\n[Security Alerts\npage](https://github.com/AzureAD/microsoft-authentication-library-for-js/network/alerts).\n\n</details>\n\nSigned-off-by: dependabot[bot] <support@github.com>\nCo-authored-by: dependabot[bot] <49699333+dependabot[bot]@users.noreply.github.com>",
          "timestamp": "2025-09-12T15:23:28-07:00",
          "tree_id": "77223d58180f58e69d7f47c71c654a31036aaab1",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/f65f2dc0dd1315acfd008b309bb5f0a1e00dd7a7"
        },
        "date": 1757734359159,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 243573,
            "range": "±0.83%",
            "unit": "ops/sec",
            "extra": "232 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 245378,
            "range": "±0.81%",
            "unit": "ops/sec",
            "extra": "234 samples"
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
          "id": "aafeda0fe35c196706ce2c6d688eea2d312c4b9a",
          "message": "Reset Redirecturi during broker fallback (#8049)\n\nIf a user attempts a broker flow but the broker isn’t available on their\nmachine, the flow falls back to the browser. However, since a redirect\nURI was provided for the broker flow, an error is thrown.\n\nThis PR adds a check to determine whether the broker is enabled in the\nconfig and resets the redirect URI to an empty string instead of\nthrowing an error.",
          "timestamp": "2025-09-16T09:26:43-07:00",
          "tree_id": "ee5c6c1c4ba51fba7d30e2ef4dd46de6318d73ee",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/aafeda0fe35c196706ce2c6d688eea2d312c4b9a"
        },
        "date": 1758057926922,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 251102,
            "range": "±0.94%",
            "unit": "ops/sec",
            "extra": "222 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 248969,
            "range": "±0.74%",
            "unit": "ops/sec",
            "extra": "233 samples"
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
          "id": "938496aa336c720ef06771252419b80ecabe421d",
          "message": "Instrument data boundary claim (#8054)\n\nThis PR adds support for the data boundary claim (xms_tdbr) from\nMicrosoft Identity Platform to the MSAL.js authentication libraries. The\ndata boundary claim indicates the geographic location where user data is\nstored and processed, particularly relevant for non-US tenants that\noperate within specific data residency boundaries.\n\n- Add dataBoundary property to AccountInfo and AccountEntity types\n- Extract data boundary information from the xms_tdbr field in client\ninfo during account creation\n- Instrument data boundary value in performance telemetry for monitoring\npurposes",
          "timestamp": "2025-09-17T15:47:11-04:00",
          "tree_id": "ca2f15aeb6be46224ffc388179c75f7aa50fbe74",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/938496aa336c720ef06771252419b80ecabe421d"
        },
        "date": 1758147328894,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 256459,
            "range": "±0.59%",
            "unit": "ops/sec",
            "extra": "236 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 257085,
            "range": "±0.64%",
            "unit": "ops/sec",
            "extra": "236 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "45603194+shenj@users.noreply.github.com",
            "name": "Jian Shen",
            "username": "shenj"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "71fa65134ade574e126eb70f9824621bf3f63ab1",
          "message": "Add changes in network and interaction client layers to support MFA (#8048)",
          "timestamp": "2025-09-18T21:10:59+01:00",
          "tree_id": "478b0558f7c732c0dcaaea7cc57e9da973dfa91b",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/71fa65134ade574e126eb70f9824621bf3f63ab1"
        },
        "date": 1758231711129,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 251033,
            "range": "±0.75%",
            "unit": "ops/sec",
            "extra": "227 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 250006,
            "range": "±0.68%",
            "unit": "ops/sec",
            "extra": "224 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "45603194+shenj@users.noreply.github.com",
            "name": "Jian Shen",
            "username": "shenj"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "4ab3202e1e4f584f092a0f0b345740395d350e88",
          "message": "Add 2 more Dublin teammates as owners for the native auth feature (#8058)",
          "timestamp": "2025-09-20T05:53:38-07:00",
          "tree_id": "111fd5906da519f093215ceba386b1fecd62b6f3",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/4ab3202e1e4f584f092a0f0b345740395d350e88"
        },
        "date": 1758381256069,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 243877,
            "range": "±0.77%",
            "unit": "ops/sec",
            "extra": "213 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 243241,
            "range": "±0.87%",
            "unit": "ops/sec",
            "extra": "220 samples"
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
          "id": "579f77570d64bd7a81b9986c8de54735edf28921",
          "message": "Release PR: official (#8057)\n\nThis PR contains the changelogs and version bumps for the MSAL.js 3P\nreleases.\n\nCo-authored-by: MSAL.js Release Automation <msaljsbuilds@microsoft.com>\nCo-authored-by: Sameera Gajjarapu <sameera.gajjarapu@microsoft.com>",
          "timestamp": "2025-09-22T08:43:51-07:00",
          "tree_id": "a0af4cb6ff0af368b05738c22e43097e10925061",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/579f77570d64bd7a81b9986c8de54735edf28921"
        },
        "date": 1758581041503,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 245034,
            "range": "±0.81%",
            "unit": "ops/sec",
            "extra": "234 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 250781,
            "range": "±0.65%",
            "unit": "ops/sec",
            "extra": "235 samples"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "45603194+shenj@users.noreply.github.com",
            "name": "Jian Shen",
            "username": "shenj"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "eb0997dda46442a5bba6bdee512c47abe3b019d0",
          "message": "[Native Auth] Add MFA related states and results (#8060)",
          "timestamp": "2025-09-24T12:56:14+01:00",
          "tree_id": "a5f46998422a93c1f6e03660b99fe09c17493e60",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/eb0997dda46442a5bba6bdee512c47abe3b019d0"
        },
        "date": 1758726882364,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 243517,
            "range": "±0.77%",
            "unit": "ops/sec",
            "extra": "232 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 241761,
            "range": "±0.81%",
            "unit": "ops/sec",
            "extra": "232 samples"
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
          "id": "5ecdba1412ead945863f5f8df308b70deb14d37f",
          "message": "Add support for \"pageException\" error to fallback to web (#8064)\n\n`PageException` is currently not being handled as a fatal error by MSAL\nJS, but the server falls back to web if it encounters the same. This PR\naddresses the gap.",
          "timestamp": "2025-09-25T11:31:06-07:00",
          "tree_id": "83ccf183c8c6472c733346a4b4867c4f9b208f63",
          "url": "https://github.com/ehtick/microsoft-authentication-library-for-js/commit/5ecdba1412ead945863f5f8df308b70deb14d37f"
        },
        "date": 1758834880912,
        "tool": "benchmarkjs",
        "benches": [
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsFirstItemInTheCache",
            "value": 256372,
            "range": "±0.72%",
            "unit": "ops/sec",
            "extra": "223 samples"
          },
          {
            "name": "ConfidentialClientApplication#acquireTokenByClientCredential-fromCache-resourceIsLastItemInTheCache",
            "value": 250258,
            "range": "±0.94%",
            "unit": "ops/sec",
            "extra": "236 samples"
          }
        ]
      }
    ]
  }
}