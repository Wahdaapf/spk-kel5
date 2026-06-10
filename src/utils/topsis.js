// src/utils/topsis.js

/**
 * Menghitung perankingan menggunakan metode TOPSIS
 * @param {number[][]} X - Matriks keputusan awal (hanya angka)
 * @param {number[]} W - Array bobot kriteria (dari CRITIC/Entropy)
 * @param {string[]} types - Array "Benefit" atau "Cost"
 */
export function calculateTopsis(X, W, types) {
    if (!X || X.length === 0 || !W || W.length === 0) return null;

    const m = X.length;
    const n = X[0].length;

    // 1. MATRIKS NORMALISASI (R) -> TOPSIS pakai pembagi Akar Kuadrat
    let norm = Array.from({ length: m }, () => Array(n).fill(0));
    let sumSquares = Array(n).fill(0);

    for (let j = 0; j < n; j++) {
        for (let i = 0; i < m; i++) {
            sumSquares[j] += Math.pow(X[i][j], 2);
        }
        sumSquares[j] = Math.sqrt(sumSquares[j]);
    }

    for (let j = 0; j < n; j++) {
        let divisor = sumSquares[j] === 0 ? 1 : sumSquares[j];
        for (let i = 0; i < m; i++) {
            norm[i][j] = X[i][j] / divisor;
        }
    }

    // 2. MATRIKS TERNORMALISASI TERBOBOT (Y)
    let Y = Array.from({ length: m }, () => Array(n).fill(0));
    for (let i = 0; i < m; i++) {
        for (let j = 0; j < n; j++) {
            Y[i][j] = norm[i][j] * W[j];
        }
    }

    // 3. SOLUSI IDEAL POSITIF (A+) DAN NEGATIF (A-)
    let A_plus = Array(n).fill(0);
    let A_min = Array(n).fill(0);

    for (let j = 0; j < n; j++) {
        let colValues = Y.map(row => row[j]);
        let maxVal = Math.max(...colValues);
        let minVal = Math.min(...colValues);

        if (types[j] === "Cost") {
            A_plus[j] = minVal;
            A_min[j] = maxVal;
        } else { // Benefit
            A_plus[j] = maxVal;
            A_min[j] = minVal;
        }
    }

    // 4. JARAK DARI SOLUSI IDEAL (D+ dan D-)
    let D_plus = Array(m).fill(0);
    let D_min = Array(m).fill(0);

    for (let i = 0; i < m; i++) {
        let sumPlus = 0;
        let sumMin = 0;
        for (let j = 0; j < n; j++) {
            sumPlus += Math.pow(Y[i][j] - A_plus[j], 2);
            sumMin += Math.pow(Y[i][j] - A_min[j], 2);
        }
        D_plus[i] = Math.sqrt(sumPlus);
        D_min[i] = Math.sqrt(sumMin);
    }

    // 5. NILAI PREFERENSI (V) DAN RANKING
    // V_i = D-_i / (D-_i + D+_i)
    let preferences = Array(m).fill(0);
    for (let i = 0; i < m; i++) {
        let totalD = D_min[i] + D_plus[i];
        preferences[i] = totalD === 0 ? 0 : D_min[i] / totalD;
    }

    // Buat array object untuk di-sorting (agar kita tahu alternatif mana yang juara)
    let rankingObj = preferences.map((val, index) => ({
        originalIndex: index,
        score: val
    }));

    // Sort Descending (Tertinggi ke Terendah)
    rankingObj.sort((a, b) => b.score - a.score);

    // Buat array hasil ranking final (Rank 1, 2, 3...) untuk masing-masing alternatif
    let finalRank = Array(m).fill(0);
    rankingObj.forEach((item, rankIndex) => {
        finalRank[item.originalIndex] = rankIndex + 1;
    });

    return { norm, Y, A_plus, A_min, D_plus, D_min, preferences, finalRank };
}