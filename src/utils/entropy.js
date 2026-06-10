// src/utils/entropy.js

export function calculateEntropy(X) {
    if (!X || X.length === 0 || !X[0]) return null;

    const m = X.length;
    const n = X[0].length;

    // 1. NORMALISASI (sama seperti Excel)
    const norm = Array.from(
        { length: m },
        () => Array(n).fill(0)
    );

    const maxX = Array(n).fill(-Infinity);

    for (let j = 0; j < n; j++) {
        for (let i = 0; i < m; i++) {
            maxX[j] = Math.max(maxX[j], X[i][j]);
        }
    }

    for (let j = 0; j < n; j++) {
        for (let i = 0; i < m; i++) {
            norm[i][j] =
                maxX[j] === 0
                    ? 0
                    : X[i][j] / maxX[j];
        }
    }

    // 2. MATRIKS PROPORSI
    const proportion = Array.from(
        { length: m },
        () => Array(n).fill(0)
    );

    for (let j = 0; j < n; j++) {
        const colSum = norm.reduce(
            (sum, row) => sum + row[j],
            0
        );

        for (let i = 0; i < m; i++) {
            proportion[i][j] =
                colSum === 0
                    ? 0
                    : norm[i][j] / colSum;
        }
    }

    // 3. ENTROPY
    const k = 1 / Math.log(m);

    const entropy = Array(n).fill(0);

    for (let j = 0; j < n; j++) {
        let sum = 0;

        for (let i = 0; i < m; i++) {
            const p = proportion[i][j];

            if (p > 0) {
                sum += p * Math.log(p);
            }
        }

        entropy[j] = -k * sum;
    }

    // 4. DISPERSI
    const dispersion = entropy.map(
        e => 1 - e
    );

    // 5. BOBOT
    const totalDispersion =
        dispersion.reduce(
            (a, b) => a + b,
            0
        );

    const weights = dispersion.map(
        d =>
            totalDispersion === 0
                ? 0
                : d / totalDispersion
    );

    return {
        norm,
        proportion,
        entropy,
        dispersion,
        weights,
    };
}