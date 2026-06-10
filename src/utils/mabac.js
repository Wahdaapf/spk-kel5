// src/utils/mabac.js

export function calculateMabac(X, W, types) {
    if (!X || !W) return null;

    const m = X.length;
    const n = X[0].length;

    // 1. NORMALISASI
    const norm = Array.from({ length: m }, () =>
        Array(n).fill(0)
    );

    const minX = Array(n).fill(Infinity);
    const maxX = Array(n).fill(-Infinity);

    for (let j = 0; j < n; j++) {
        for (let i = 0; i < m; i++) {
            minX[j] = Math.min(minX[j], X[i][j]);
            maxX[j] = Math.max(maxX[j], X[i][j]);
        }
    }

    for (let j = 0; j < n; j++) {
        const divisor = maxX[j] - minX[j] || 1;

        for (let i = 0; i < m; i++) {
            norm[i][j] =
                types[j] === "Cost"
                    ? (maxX[j] - X[i][j]) / divisor
                    : (X[i][j] - minX[j]) / divisor;
        }
    }

    // 2. MATRIKS V
    const V = Array.from({ length: m }, () =>
        Array(n).fill(0)
    );

    for (let i = 0; i < m; i++) {
        for (let j = 0; j < n; j++) {
            V[i][j] = W[j] * (norm[i][j] + 1);
        }
    }

    // 3. BORDER APPROXIMATION AREA (G)
    const G = Array(n).fill(0);

    for (let j = 0; j < n; j++) {
        let product = 1;

        for (let i = 0; i < m; i++) {
            product *= V[i][j];
        }

        G[j] = Math.pow(product, 1 / m);
    }

    // 4. Q MATRIX
    const Q = Array.from({ length: m }, () =>
        Array(n).fill(0)
    );

    for (let i = 0; i < m; i++) {
        for (let j = 0; j < n; j++) {
            Q[i][j] = V[i][j] - G[j];
        }
    }

    // 5. NILAI S
    const S = Array(m).fill(0);

    for (let i = 0; i < m; i++) {
        S[i] = Q[i].reduce(
            (sum, value) => sum + value,
            0
        );
    }

    // 6. RANKING
    const rankingObj = S.map((score, index) => ({
        index,
        score,
    }));

    rankingObj.sort(
        (a, b) => b.score - a.score
    );

    const finalRank = Array(m).fill(0);

    rankingObj.forEach((item, idx) => {
        finalRank[item.index] = idx + 1;
    });

    return {
        norm,
        V,
        G,
        Q,
        S,
        finalRank,
    };
}