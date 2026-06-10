// src/utils/critic.js

/**
 * Menghitung bobot kriteria menggunakan metode CRITIC
 * @param {number[][]} X - Matriks data (hanya angka)
 * @param {string[]} types - Array "Benefit" atau "Cost"
 */
export function calculateCritic(X, types) {
    if (!X || X.length === 0 || !X[0] || X[0].length === 0) return null;

    const m = X.length; // Jumlah Alternatif
    const n = X[0].length; // Jumlah Kriteria

    // 1. NORMALISASI MIN-MAX
    let norm = Array.from({ length: m }, () => Array(n).fill(0));
    let minX = Array(n).fill(Infinity);
    let maxX = Array(n).fill(-Infinity);

    for (let j = 0; j < n; j++) {
        for (let i = 0; i < m; i++) {
            if (X[i][j] < minX[j]) minX[j] = X[i][j];
            if (X[i][j] > maxX[j]) maxX[j] = X[i][j];
        }
    }

    for (let j = 0; j < n; j++) {
        let divisor = maxX[j] - minX[j];
        if (divisor === 0) divisor = 1;
        for (let i = 0; i < m; i++) {
            if (types[j] === "Cost") {
                norm[i][j] = (maxX[j] - X[i][j]) / divisor;
            } else {
                norm[i][j] = (X[i][j] - minX[j]) / divisor;
            }
        }
    }

    // 2. STANDAR DEVIASI (σ)
    let stdDev = Array(n).fill(0);
    for (let j = 0; j < n; j++) {
        let mean = norm.reduce((sum, row) => sum + row[j], 0) / m;
        // Standar Deviasi Populasi
        let variance =
            norm.reduce(
                (sum, row) => sum + Math.pow(row[j] - mean, 2),
                0
            ) / (m - 1);
        stdDev[j] = Math.sqrt(variance);
    }

    // 3. KORELASI & KONFLIK (C)
    let correlation = Array.from({ length: n }, () => Array(n).fill(0));
    let conflict = Array(n).fill(0);

    for (let j = 0; j < n; j++) {
        let sumConflict = 0;
        for (let k = 0; k < n; k++) {
            let meanJ = norm.reduce((sum, row) => sum + row[j], 0) / m;
            let meanK = norm.reduce((sum, row) => sum + row[k], 0) / m;

            let num = 0, denJ = 0, denK = 0;
            for (let i = 0; i < m; i++) {
                num += (norm[i][j] - meanJ) * (norm[i][k] - meanK);
                denJ += Math.pow(norm[i][j] - meanJ, 2);
                denK += Math.pow(norm[i][k] - meanK, 2);
            }

            let r_jk = 0;
            if (denJ * denK > 0) r_jk = num / Math.sqrt(denJ * denK);

            correlation[j][k] = r_jk;
            sumConflict += (1 - r_jk); // Konflik = 1 - r_jk
        }
        conflict[j] = sumConflict;
    }

    // 4. NILAI INFORMASI (I)
    let info = Array(n).fill(0);
    let totalInfo = 0;
    for (let j = 0; j < n; j++) {
        info[j] = stdDev[j] * conflict[j];
        totalInfo += info[j];
    }

    // 5. BOBOT AKHIR (W)
    let weights = info.map(i => (totalInfo > 0 ? i / totalInfo : 0));

    return { norm, stdDev, correlation, conflict, info, weights };
}