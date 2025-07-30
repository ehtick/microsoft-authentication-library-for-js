window.BENCHMARK_DATA = {
  "lastUpdate": 1753866886840,
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
      }
    ]
  }
}