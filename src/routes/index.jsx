import { createFileRoute } from "@tanstack/react-router";
import { useState, useMemo } from "react";
import * as XLSX from "xlsx";
import { calculateCritic } from "../utils/critic";
import { calculateEntropy } from "../utils/entropy";

import { calculateTopsis } from "../utils/topsis";
import { calculateMabac } from "../utils/mabac";

export const Route = createFileRoute("/")({
    head: () => ({
        meta: [
            { title: "Decidra — Sistem Pendukung Keputusan" },
            { name: "description", content: "Aplikasi pengambilan keputusan multi-kriteria dengan metode CRITIC, Entropy, TOPSIS, dan MABAC." },
            { property: "og:title", content: "Decidra — Sistem Pendukung Keputusan" },
            { property: "og:description", content: "Pengambilan keputusan multi-kriteria yang elegan dan cepat." },
        ],
    }),
    component: Index,
});

function Index() {
    const [step, setStep] = useState(1);
    const [dim, setDim] = useState({ alt: 3, crit: 3 });
    const [headers, setHeaders] = useState([]);
    const [types, setTypes] = useState([]);
    const [rows, setRows] = useState([]);
    const [method, setMethod] = useState({ bobot: "CRITIC", ranking: "TOPSIS" });
    const [activeTab, setActiveTab] = useState(0);

    const downloadTemplate = () => {
        const wsData = [];
        const headerRow = ["Alternatif"];
        for (let i = 1; i <= dim.crit; i++) headerRow.push(`Kriteria ${i}`);
        wsData.push(headerRow);
        const typeRow = ["Tipe (Isi C/B)"];
        for (let i = 1; i <= dim.crit; i++) typeRow.push("Benefit");
        wsData.push(typeRow);
        for (let i = 1; i <= dim.alt; i++) {
            const row = [`Alternatif ${i}`];
            for (let j = 1; j <= dim.crit; j++) row.push(0);
            wsData.push(row);
        }
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Template_SPK");
        XLSX.writeFile(wb, "Template_SPK.xlsx");
    };

    const handleFileUpload = (e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (evt) => {
            const bstr = evt.target?.result;
            const wb = XLSX.read(bstr, { type: "binary" });
            const wsname = wb.SheetNames[0];
            const ws = wb.Sheets[wsname];
            const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
            setHeaders(data[0] || []);
            setTypes(data[1] || []);
            setRows(data.slice(2) || []);
            setStep(2);
        };
        reader.readAsBinaryString(file);
    };

    const updateData = (rowIndex, colIndex, value) => {
        const newRows = rows.map((r) => [...r]);
        newRows[rowIndex][colIndex] = value;
        setRows(newRows);
    };

    const steps = [
        { n: 1, label: "Setup" },
        { n: 2, label: "Review Data" },
        { n: 3, label: "Kalkulasi" },
    ];

    const calculation = useMemo(() => {
        if (!rows || rows.length === 0 || !headers || headers.length <= 1) return null;

        // 1. Ekstrak data untuk perhitungan
        const altNames = rows.map((r) => r[0]);
        const critNames = headers.slice(1);
        const X = rows.map((row) => row.slice(1).map(Number));
        const critTypes = types.slice(1);

        if (X.length === 0 || X[0].length === 0) return null;

        // 2. Hitung Pembobotan
        let resultBobot = null;
        let W = [];

        if (method.bobot === "CRITIC") {
            resultBobot = calculateCritic(X, critTypes);
        }
        else if (method.bobot === "Entropy") {
            resultBobot = calculateEntropy(X, critTypes);
        }

        W = resultBobot?.weights || [];

        // 3. Hitung Perankingan
        let resultRanking = null;

        if (W.length > 0) {
            if (method.ranking === "TOPSIS") {
                resultRanking = calculateTopsis(
                    X,
                    W,
                    critTypes
                );
            }

            else if (method.ranking === "MABAC") {
                resultRanking = calculateMabac(
                    X,
                    W,
                    critTypes
                );
            }
        }

        console.log("BOBOT =", method.bobot);
        console.log("RANKING =", method.ranking);
        console.log("RESULT BOBOT =", resultBobot);
        console.log("RESULT RANKING =", resultRanking);

        return { altNames, critNames, X, critTypes, bobot: resultBobot, ranking: resultRanking, finalWeights: W };
    }, [rows, headers, types, method]);

    console.log(calculation);

    const renderMatrixTable = (matrixArray) => {
        if (!matrixArray) return <p className="text-muted-foreground p-4">Data sedang diproses...</p>;
        return (
            <div className="overflow-x-auto rounded-xl border border-border hide-scrollbar bg-card/40 mt-4">
                <table className="min-w-full text-sm text-left whitespace-nowrap">
                    <thead className="bg-white/[0.03]">
                        <tr>
                            <th className="px-5 py-3 border-b border-border font-semibold">Alternatif</th>
                            {calculation.critNames.map((c, i) => (
                                <th key={i} className="px-5 py-3 border-b border-border font-semibold">{c}</th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {matrixArray.map((row, i) => (
                            <tr key={i} className="border-b border-border/50 hover:bg-white/[0.02]">
                                <td className="px-5 py-3 font-medium bg-white/[0.01]">{calculation.altNames[i]}</td>
                                {row.map((val, j) => (
                                    <td key={j} className="px-5 py-3 tabular-nums">{Number(val).toFixed(4)}</td>
                                ))}
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
        );
    };

    const tabs = useMemo(() => {
        const result = [
            {
                key: "matrix-awal",
                title: "Matriks Keputusan Awal (X)",
                desc: "Menampilkan matriks dari data Excel yang diunggah.",
                formula: null,
                source: calculation?.X,
                type: "matrix",
            }
        ];

        if (method.bobot === "CRITIC") {
            result.push(
                {
                    key: "critic-norm",
                    title: "Matriks Normalisasi",
                    desc: "Normalisasi Min-Max berdasarkan tipe Benefit dan Cost.",
                    formula: "r_ij=(x_ij-min)/(max-min)",
                    source: calculation?.bobot?.norm,
                    type: "matrix",
                },
                {
                    key: "critic-stddev",
                    title: "Standar Deviasi",
                    desc: "Menghitung penyebaran nilai hasil normalisasi.",
                    formula: "σ_j=STDEV.S(r_j)",
                    source: calculation?.bobot?.stdDev,
                    type: "vector",
                },
                {
                    key: "critic-correlation",
                    title: "Korelasi",
                    desc: "Mengukur hubungan antar kriteria.",
                    formula: "r_jk",
                    source: calculation?.bobot?.correlation,
                    type: "matrix",
                },
                {
                    key: "critic-conflict",
                    title: "Konflik",
                    desc: "Σ(1-rjk)",
                    formula: "C_j=Σ(1-r_jk)",
                    source: calculation?.bobot?.conflict,
                    type: "vector",
                },
                {
                    key: "critic-info",
                    title: "Nilai Informasi",
                    desc: "Informasi tiap kriteria.",
                    formula: "I_j=σ_j×C_j",
                    source: calculation?.bobot?.info,
                    type: "vector",
                },
                {
                    key: "critic-weight",
                    title: "Bobot CRITIC",
                    desc: "Normalisasi nilai informasi.",
                    formula: "w_j=I_j/ΣI_j",
                    source: calculation?.bobot?.weights,
                    type: "vector",
                }
            );
        }

        if (method.bobot === "Entropy") {
            result.push(
                {
                    key: "entropy-norm",
                    title: "Normalisasi (Entropy)",
                    desc: "Normalisasi data",
                    formula: "(x-min)/(max-min)",
                    source: calculation?.bobot?.norm,
                    type: "matrix",
                },
                {
                    key: "entropy-proportion",
                    title: "Matriks Proporsi",
                    desc: "Pij = rij / Σrij",
                    formula: "p_ij=r_ij/Σr_ij",
                    source: calculation?.bobot?.proportion,
                    type: "matrix",
                },
                {
                    key: "entropy-value",
                    title: "Nilai Entropy",
                    desc: "Entropy tiap kriteria",
                    formula: "E_j=-kΣ(p_ij ln p_ij)",
                    source: calculation?.bobot?.entropy,
                    type: "vector",
                },
                {
                    key: "entropy-dispersion",
                    title: "Dispersi",
                    desc: "Dj = 1 - Ej",
                    formula: "D_j=1-E_j",
                    source: calculation?.bobot?.dispersion,
                    type: "vector",
                },
                {
                    key: "entropy-weight",
                    title: "Bobot Entropi",
                    desc: "Normalisasi dispersi",
                    formula: "w_j=D_j/ΣD_j",
                    source: calculation?.bobot?.weights,
                    type: "vector",
                },
            )
        }

        if (method.ranking === "TOPSIS") {
            result.push(
                {
                    key: "topsis-norm",
                    title: "Matriks Normalisasi TOPSIS",
                    desc: "Normalisasi vektor.",
                    formula: "r_ij=x_ij/√Σx²",
                    source: calculation?.ranking?.norm,
                    type: "matrix",
                },
                {
                    key: "topsis-y",
                    title: "Matriks Y Terbobot",
                    desc: "Normalisasi × bobot.",
                    formula: "y_ij=r_ij×w_j",
                    source: calculation?.ranking?.Y,
                    type: "matrix",
                },
                {
                    key: "topsis-ideal",
                    title: "A+ & A-",
                    desc: "Solusi ideal.",
                    formula: null,
                    source: {
                        A_plus: calculation?.ranking?.A_plus,
                        A_min: calculation?.ranking?.A_min,
                    },
                    type: "object",
                },
                {
                    key: "topsis-distance",
                    title: "D+ & D-",
                    desc: "Jarak solusi ideal.",
                    formula: null,
                    source: {
                        D_plus: calculation?.ranking?.D_plus,
                        D_min: calculation?.ranking?.D_min,
                    },
                    type: "object",
                },
                {
                    key: "topsis-ranking",
                    title: "Preferensi & Ranking",
                    desc: "Hasil akhir TOPSIS.",
                    formula: "V_i=D^-/(D^++D^-)",
                    source: {
                        preferences: calculation?.ranking?.preferences,
                        finalRank: calculation?.ranking?.finalRank,
                    },
                    type: "object",
                }
            );
        }

        if (method.ranking === "MABAC") {
            result.push(
                {
                    key: "mabac-norm",
                    title: "Normalisasi (MABAC)",
                    desc: "Normalisasi matriks",
                    formula: "(x-min)/(max-min)",
                    source: calculation?.ranking?.norm,
                    type: "matrix",
                },
                {
                    key: "mabac-v",
                    title: "Matriks V",
                    desc: "Matriks terbobot",
                    formula: "v_ij=w_j(r_ij+1)",
                    source: calculation?.ranking?.V,
                    type: "matrix",
                },
                {
                    key: "mabac-g",
                    title: "Border Approx (G)",
                    desc: "Geometric Mean",
                    formula: "G_j=(Πv_ij)^(1/m)",
                    source: calculation?.ranking?.G,
                    type: "vector",
                },
                {
                    key: "mabac-q",
                    title: "Q Matrix",
                    desc: "Q = V - G",
                    formula: "q_ij=v_ij-G_j",
                    source: calculation?.ranking?.Q,
                    type: "matrix",
                },
                {
                    key: "mabac-ranking",
                    title: "Nilai Akhir S & Ranking",
                    desc: "Perankingan MABAC",
                    formula: "S_i=Σq_ij",
                    source: {
                        S: calculation?.ranking?.S,
                        finalRank: calculation?.ranking?.finalRank,
                    },
                    type: "object",
                },
            )
        }

        return result;
    }, [calculation, method]);

    console.log(method);
    console.log(tabs.map(t => t.title));
    const currentTab = tabs[activeTab];

    const renderTabContent = (tab) => {
        if (!tab) return null;

        switch (tab.type) {
            case "matrix":
                return renderMatrixTable(tab.source);

            case "vector":
                return (
                    <div className="overflow-x-auto rounded-xl border border-border bg-card/40 mt-4">
                        <table className="min-w-full text-sm text-left">
                            <thead className="bg-white/[0.03]">
                                <tr>
                                    <th className="px-5 py-3 border-b border-border font-semibold">
                                        Kriteria
                                    </th>
                                    <th className="px-5 py-3 border-b border-border font-semibold">
                                        Nilai
                                    </th>
                                </tr>
                            </thead>

                            <tbody>
                                {tab.source?.map((value, i) => (
                                    <tr
                                        key={i}
                                        className="border-b border-border/50 hover:bg-white/[0.02]"
                                    >
                                        <td className="px-5 py-3 font-medium">
                                            {calculation?.critNames?.[i]}
                                        </td>

                                        <td className="px-5 py-3 tabular-nums">
                                            {typeof value === "number"
                                                ? value.toFixed(6)
                                                : value}
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                );

            case "object":
                if (tab.key === "topsis-ideal") {
                    return (
                        <div className="overflow-x-auto rounded-xl border border-border bg-card/40 mt-4">
                            <table className="min-w-full text-sm text-left">
                                <thead className="bg-white/[0.03]">
                                    <tr>
                                        <th className="px-5 py-3 border-b border-border font-semibold">
                                            Kriteria
                                        </th>

                                        <th className="px-5 py-3 border-b border-border font-semibold">
                                            A+
                                        </th>

                                        <th className="px-5 py-3 border-b border-border font-semibold">
                                            A-
                                        </th>
                                    </tr>
                                </thead>

                                <tbody>
                                    {calculation.critNames.map((crit, i) => (
                                        <tr
                                            key={i}
                                            className="border-b border-border/50 hover:bg-white/[0.02]"
                                        >
                                            <td className="px-5 py-3 font-medium">
                                                {crit}
                                            </td>

                                            <td className="px-5 py-3">
                                                {tab.source.A_plus[i].toFixed(6)}
                                            </td>

                                            <td className="px-5 py-3">
                                                {tab.source.A_min[i].toFixed(6)}
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    );
                }

                if (tab.key === "topsis-distance") {
                    return (
                        <div className="overflow-x-auto rounded-xl border border-border bg-card/40 mt-4">
                            <table className="min-w-full text-sm text-left whitespace-nowrap">
                                <thead className="bg-white/[0.03]">
                                    <tr>
                                        <th className="px-5 py-3 border-b border-border font-semibold">
                                            Alternatif
                                        </th>

                                        <th className="px-5 py-3 border-b border-border font-semibold">
                                            D+
                                        </th>

                                        <th className="px-5 py-3 border-b border-border font-semibold">
                                            D-
                                        </th>
                                    </tr>
                                </thead>

                                <tbody>
                                    {calculation.altNames.map((alt, i) => (
                                        <tr
                                            key={i}
                                            className="border-b border-border/50 hover:bg-white/[0.02]"
                                        >
                                            <td className="px-5 py-3 font-medium">
                                                {alt}
                                            </td>

                                            <td className="px-5 py-3">
                                                {tab.source.D_plus[i].toFixed(6)}
                                            </td>

                                            <td className="px-5 py-3">
                                                {tab.source.D_min[i].toFixed(6)}
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    );
                }

                if (tab.key === "topsis-ranking") {
                    return (
                        <div className="overflow-x-auto rounded-xl border border-border bg-card/40 mt-4">
                            <table className="min-w-full text-sm text-left whitespace-nowrap">
                                <thead className="bg-white/[0.03]">
                                    <tr>
                                        <th className="px-5 py-3 border-b border-border font-semibold">
                                            Alternatif
                                        </th>

                                        <th className="px-5 py-3 border-b border-border font-semibold">
                                            Preferensi
                                        </th>

                                        <th className="px-5 py-3 border-b border-border font-semibold">
                                            Ranking
                                        </th>
                                    </tr>
                                </thead>

                                <tbody>
                                    {calculation.altNames.map((alt, i) => (
                                        <tr
                                            key={i}
                                            className="border-b border-border/50 hover:bg-white/[0.02]"
                                        >
                                            <td className="px-5 py-3 font-medium">
                                                {alt}
                                            </td>

                                            <td className="px-5 py-3">
                                                {tab.source.preferences[i].toFixed(6)}
                                            </td>

                                            <td className="px-5 py-3">
                                                {tab.source.finalRank[i]}
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    );
                }

                if (tab.key === "mabac-ranking") {
                    return (
                        <div className="overflow-x-auto rounded-xl border border-border bg-card/40 mt-4">
                            <table className="min-w-full text-sm text-left whitespace-nowrap">
                                <thead className="bg-white/[0.03]">
                                    <tr>
                                        <th className="px-5 py-3 border-b border-border font-semibold">
                                            Alternatif
                                        </th>

                                        <th className="px-5 py-3 border-b border-border font-semibold">
                                            Nilai S
                                        </th>

                                        <th className="px-5 py-3 border-b border-border font-semibold">
                                            Ranking
                                        </th>
                                    </tr>
                                </thead>

                                <tbody>
                                    {calculation.altNames.map((alt, i) => (
                                        <tr
                                            key={i}
                                            className="border-b border-border/50 hover:bg-white/[0.02]"
                                        >
                                            <td className="px-5 py-3 font-medium">
                                                {alt}
                                            </td>

                                            <td className="px-5 py-3">
                                                {tab.source.S[i].toFixed(6)}
                                            </td>

                                            <td className="px-5 py-3">
                                                {tab.source.finalRank[i]}
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    );
                }

            default:
                return null;
        }
    };

    return (
        <div className="relative min-h-screen overflow-hidden">
            {/* Yellow ambient backdrop orbs */}
            <div className="pointer-events-none absolute inset-0 -z-10">
                <div className="absolute -top-40 -left-20 h-[480px] w-[480px] rounded-full bg-gold/20 blur-3xl animate-float-slow" />
                <div className="absolute top-1/3 -right-32 h-[520px] w-[520px] rounded-full bg-yellow-glow/18 blur-3xl animate-float-slow-rev" />
                <div className="absolute bottom-0 left-1/3 h-[360px] w-[360px] rounded-full bg-yellow-bright/14 blur-3xl animate-float-slow" />
                <div className="absolute top-1/2 left-10 h-[300px] w-[300px] rounded-full bg-gold/12 blur-3xl animate-float-slow-rev" />
            </div>

            <div className="mx-auto max-w-6xl px-6 py-16">
                {/* Header */}
                <header className="mb-10 animate-pop-in">

                    <h1 className="font-display text-5xl md:text-6xl font-bold leading-[1.05]">
                        <span className="text-glow">Sistem Pendukung Keputusan</span>
                    </h1>
                    <p className="mt-4 max-w-2xl text-muted-foreground text-lg">
                        Analisis alternatif Anda dengan CRITIC, Entropy, TOPSIS, dan MABAC.
                    </p>
                </header>

                {/* Stepper */}
                <nav className="mb-8 flex items-center gap-2 animate-slide-up">
                    {steps.map((s, i) => (
                        <div key={s.n} className="flex items-center gap-2">
                            <button
                                onClick={() => s.n < step && setStep(s.n)}
                                className={`flex items-center gap-2 rounded-full px-4 py-2 text-sm font-medium transition-all duration-300 ${step === s.n
                                    ? "bg-gradient-primary text-background shadow-lg glow-yellow scale-105"
                                    : step > s.n
                                        ? "bg-accent/20 text-primary hover:bg-accent/30"
                                        : "bg-card/60 text-muted-foreground"
                                    }`}
                            >
                                <span className={`flex h-5 w-5 items-center justify-center rounded-full text-xs font-bold ${step >= s.n ? "bg-background/20" : "bg-muted"}`}>
                                    {step > s.n ? "✓" : s.n}
                                </span>
                                {s.label}
                            </button>
                            {i < steps.length - 1 && (
                                <div className={`h-px w-8 transition-colors ${step > s.n ? "bg-accent" : "bg-border"}`} />
                            )}
                        </div>
                    ))}
                </nav>

                {/* STEP 1 */}
                {step === 1 && (
                    <section className="glass rounded-3xl p-8 md:p-10 animate-pop-in">
                        <div className="mb-6">
                            <h2 className="font-display text-2xl font-bold mb-1">Langkah 1 · Setup Dimensi</h2>
                            <p className="text-muted-foreground">Tentukan jumlah alternatif dan kriteria, lalu unduh template.</p>
                        </div>

                        <div className="grid sm:grid-cols-2 gap-5 mb-8">
                            <NumberField label="Jumlah Alternatif" value={dim.alt} onChange={(v) => setDim({ ...dim, alt: v })} />
                            <NumberField label="Jumlah Kriteria" value={dim.crit} onChange={(v) => setDim({ ...dim, crit: v })} />
                        </div>

                        <div className="flex flex-wrap gap-3">
                            <button
                                onClick={downloadTemplate}
                                className="group relative overflow-hidden rounded-xl bg-gradient-gold px-6 py-3 font-semibold text-background transition-transform hover:scale-[1.03] active:scale-[0.98]"
                            >
                                <span className="relative z-10 flex items-center gap-2">⬇ Download Template Excel</span>
                            </button>
                            <label className="group relative cursor-pointer overflow-hidden rounded-xl bg-gradient-primary px-6 py-3 font-semibold text-background transition-transform hover:scale-[1.03] active:scale-[0.98] glow-yellow">
                                <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} className="absolute inset-0 cursor-pointer opacity-0" />
                                <span className="relative z-10 flex items-center gap-2">📂 Import Template Terisi</span>
                            </label>
                        </div>
                    </section>
                )}

                {/* STEP 2 */}
                {step === 2 && (
                    <section className="glass rounded-3xl p-8 md:p-10 animate-pop-in">
                        <div className="mb-6 flex items-start justify-between gap-4 flex-wrap">
                            <div>
                                <h2 className="font-display text-2xl font-bold mb-1">Langkah 2 · Review & Edit Data</h2>
                                <p className="text-muted-foreground">Periksa nilai, atur tipe (Benefit/Cost), dan pilih metode.</p>
                            </div>
                        </div>

                        <div className="overflow-x-auto rounded-2xl border border-border hide-scrollbar bg-card/40 mb-6">
                            <table className="min-w-full text-sm">
                                <thead>
                                    <tr className="bg-white/[0.03]">
                                        {headers.map((h, i) => (
                                            <th key={i} className="px-4 py-3 text-left font-semibold text-foreground border-b border-border">{h}</th>
                                        ))}
                                    </tr>
                                    <tr className="bg-white/[0.02]">
                                        {types.map((t, i) => (
                                            <th key={i} className="px-4 py-2 border-b border-border">
                                                {i === 0 ? (
                                                    <span className="text-xs italic text-muted-foreground">{t}</span>
                                                ) : (
                                                    <select
                                                        value={t}
                                                        onChange={(e) => {
                                                            const nt = [...types];
                                                            nt[i] = e.target.value;
                                                            setTypes(nt);
                                                        }}
                                                        className="w-full rounded-lg border border-border bg-background/50 px-2 py-1 text-xs font-medium outline-none focus:border-primary"
                                                    >
                                                        <option value="Benefit">Benefit</option>
                                                        <option value="Cost">Cost</option>
                                                    </select>
                                                )}
                                            </th>
                                        ))}
                                    </tr>
                                </thead>
                                <tbody>
                                    {rows.map((row, rIndex) => (
                                        <tr key={rIndex} className="transition-colors hover:bg-white/[0.04]">
                                            {row.map((cell, cIndex) => (
                                                <td key={cIndex} className="border-b border-border/60 p-0">
                                                    {cIndex === 0 ? (
                                                        <input
                                                            type="text"
                                                            value={cell}
                                                            onChange={(e) => updateData(rIndex, cIndex, e.target.value)}
                                                            className="w-full bg-transparent px-4 py-2.5 outline-none focus:bg-primary/10 transition-colors"
                                                        />
                                                    ) : (
                                                        <input
                                                            type="number"
                                                            value={cell}
                                                            onChange={(e) => updateData(rIndex, cIndex, Number(e.target.value))}
                                                            className="w-full bg-transparent px-4 py-2.5 text-right tabular-nums outline-none focus:bg-primary/10 transition-colors"
                                                        />
                                                    )}
                                                </td>
                                            ))}
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>

                        <div className="grid md:grid-cols-2 gap-5 mb-8">
                            <MethodPicker
                                label="Metode Pembobotan"
                                hint="Cara menghitung bobot tiap kriteria"
                                value={method.bobot}
                                onChange={(v) => setMethod({ ...method, bobot: v })}
                                options={[
                                    { value: "CRITIC", title: "CRITIC", desc: "Memperhitungkan kontras & korelasi antar kriteria." },
                                    { value: "Entropy", title: "Entropy", desc: "Bobot dari tingkat keragaman informasi data." },
                                ]}
                            />
                            <MethodPicker
                                label="Metode Perankingan"
                                hint="Cara mengurutkan alternatif"
                                value={method.ranking}
                                onChange={(v) => setMethod({ ...method, ranking: v })}
                                options={[
                                    { value: "TOPSIS", title: "TOPSIS", desc: "Kedekatan ke solusi ideal positif & negatif." },
                                    { value: "MABAC", title: "MABAC", desc: "Jarak dari batas area aproksimasi rerata." },
                                ]}
                            />
                        </div>

                        <div className="flex justify-between items-center">
                            <button onClick={() => setStep(1)} className="text-muted-foreground hover:text-foreground transition-colors text-sm font-medium">
                                ← Kembali
                            </button>
                            <button
                                onClick={() => setStep(3)}
                                className="rounded-xl bg-gradient-primary px-7 py-3 font-bold text-background glow-yellow transition-transform hover:scale-[1.05] active:scale-[0.97]"
                            >
                                🚀 Kerjakan Kalkulasi
                            </button>
                        </div>
                    </section>
                )}

                {/* STEP 3 */}
                {step === 3 && (
                    <section className="glass rounded-3xl p-8 md:p-10 animate-pop-in">
                        <div className="mb-6 flex justify-between items-center flex-wrap gap-3">
                            <div>
                                <h2 className="font-display text-2xl font-bold">Langkah 3 · Proses Matematis</h2>
                                <p className="text-muted-foreground text-sm mt-1">Telusuri tiap tahap kalkulasi secara terstruktur.</p>
                            </div>
                            <button onClick={() => setStep(2)} className="rounded-lg border border-border bg-card px-3 py-1.5 text-sm hover:bg-card/80 transition-colors">
                                Kembali Edit
                            </button>
                        </div>

                        <div className="flex gap-1 border-b border-border mb-6 overflow-x-auto hide-scrollbar pb-px">
                            {tabs.map((tab, index) => (
                                <button
                                    key={tab.key}
                                    onClick={() => setActiveTab(index)}
                                    className={`relative whitespace-nowrap px-4 py-2.5 text-sm font-medium transition-colors ${activeTab === index
                                        ? "text-foreground"
                                        : "text-muted-foreground hover:text-foreground"
                                        }`}
                                >
                                    {tab.title}
                                    {activeTab === index && (
                                        <span className="absolute bottom-[-1px] left-0 right-0 h-0.5 bg-gradient-primary rounded-full animate-slide-up" />
                                    )}
                                </button>
                            ))}
                        </div>

                        <div
                            key={currentTab?.key}
                            className="rounded-2xl border border-border bg-card/40 p-6 min-h-[320px] animate-pop-in"
                        >
                            <TabPane
                                title={currentTab?.title}
                                desc={currentTab?.desc}
                                formula={currentTab?.formula}
                            />

                            {renderTabContent(currentTab)}
                        </div>
                    </section>
                )}

                <footer className="mt-12 text-center text-xs text-muted-foreground">
                    Decidra · Crafted for clear decisions
                </footer>
            </div>
        </div >
    );
}

function NumberField({ label, value, onChange }) {
    return (
        <label className="block">
            <span className="block text-xs font-semibold uppercase tracking-wider text-muted-foreground mb-2">{label}</span>
            <input
                type="number"
                min={2}
                value={value}
                onChange={(e) => onChange(Number(e.target.value))}
                className="w-full rounded-xl border border-border bg-card/60 px-4 py-3 text-lg font-semibold text-foreground outline-none transition-all focus:border-primary focus:bg-card focus:ring-2 focus:ring-primary/30"
            />
        </label>
    );
}

function MethodPicker({ label, hint, value, onChange, options }) {
    return (
        <div className="rounded-2xl border border-border bg-card/40 p-5">
            <div className="mb-3">
                <div className="text-xs font-semibold uppercase tracking-wider text-primary">{label}</div>
                <div className="text-sm text-muted-foreground">{hint}</div>
            </div>
            <div className="grid grid-cols-2 gap-3">
                {options.map((o) => {
                    const active = value === o.value;
                    return (
                        <button
                            key={o.value}
                            type="button"
                            onClick={() => onChange(o.value)}
                            className={`group relative text-left rounded-xl border p-4 transition-all duration-300 ${active
                                ? "border-primary bg-primary/15 glow-yellow scale-[1.02]"
                                : "border-border bg-background/40 hover:border-primary/50 hover:bg-background/60 hover:-translate-y-0.5"
                                }`}
                        >
                            <div className="flex items-center justify-between mb-1">
                                <span className={`font-bold tracking-tight ${active ? "text-foreground" : "text-foreground/90"}`}>{o.title}</span>
                                <span
                                    className={`flex h-5 w-5 items-center justify-center rounded-full border text-[10px] font-bold transition-all ${active ? "border-primary bg-gradient-primary text-background" : "border-border text-transparent"
                                        }`}
                                >
                                    ✓
                                </span>
                            </div>
                            <p className="text-xs leading-relaxed text-muted-foreground">{o.desc}</p>
                        </button>
                    );
                })}
            </div>
        </div>
    );
}

function TabPane({ title, desc, formula }) {
    return (
        <div>
            <h3 className="font-display text-xl font-bold mb-2">{title}</h3>
            <p className="text-muted-foreground">{desc}</p>
            {formula && (
                <div className="mt-4 inline-block rounded-lg border border-border bg-background/60 px-4 py-2 font-mono text-sm text-primary">
                    {formula}
                </div>
            )}
        </div>
    );
}
