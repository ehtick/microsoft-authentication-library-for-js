window.BENCHMARK_DATA = {
  "lastUpdate": 1752551436310,
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
      }
    ]
  }
}