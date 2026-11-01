/*
 * Gantt PQR - transformation du DataView "table" en lignes Site > Produit.
 * Module pur (aucun accès DOM / réseau) : testable hors Power BI.
 */
"use strict";

/** Une réunion = une ligne de données valide. */
export interface Meeting {
    rowIndex: number;
    site: string;
    product: string;
    deliverable: string;
    start: Date;
    end: Date;
    milestone: Date | null;
    /** Progression normalisée dans [0, 1] ; null si absente. */
    progress: number | null;
    /** Valeur brute de progression (pour l'infobulle). */
    rawProgress: number | null;
    selectionId?: any;
}

export type GroupField = "site" | "product";

/**
 * Ordre des niveaux du bandeau de gauche (au-dessus des réunions) :
 * "site-product" (défaut), "product-site", "site" ou "product".
 */
export type HierarchyOrder = "site-product" | "product-site" | "site" | "product";

export const HIERARCHY_ORDERS: Array<{ value: HierarchyOrder; displayName: string }> = [
    { value: "site-product", displayName: "Site > Produit > Réunion" },
    { value: "product-site", displayName: "Produit > Site > Réunion" },
    { value: "site", displayName: "Site > Réunion" },
    { value: "product", displayName: "Produit > Réunion" }
];

export function parseHierarchyOrder(value: unknown): HierarchyOrder {
    const v = String(value);
    return HIERARCHY_ORDERS.some(h => h.value === v) ? v as HierarchyOrder : "site-product";
}

/** Ligne de regroupement (Site ou Produit, à n'importe quel niveau). */
export interface GroupRow {
    kind: "group";
    key: string;
    /** 0 = premier parent. */
    depth: number;
    field: GroupField;
    label: string;
    /** Site commun à toutes les réunions du groupe (couleur), sinon null. */
    site: string | null;
    /** Toutes les réunions du groupe (descendants compris). */
    meetings: Meeting[];
    /** Dernier niveau de regroupement (juste au-dessus des réunions). */
    leaf: boolean;
    collapsible: boolean;
    collapsed: boolean;
    start: Date | null;
    end: Date | null;
}

/** Une ligne par réunion (mode "rows"). */
export interface MeetingRow {
    kind: "meeting";
    key: string;
    depth: number;
    meeting: Meeting;
}

export type GanttRow = GroupRow | MeetingRow;

/** "rows" : une barre par ligne ; "compact" : réunions sur la ligne du dernier regroupement. */
export type DisplayMode = "rows" | "compact";

export interface RawColumns {
    site: number;
    product: number;
    deliverable: number;
    start: number;
    end: number;
    milestone: number;
    progress: number;
    siteOrder: number;
}

export const EMPTY_LABEL = "(Vide)";

/** Supprime les espaces de début/fin ; valeurs nulles -> "". */
export function cleanText(value: unknown): string {
    if (value === null || value === undefined) {
        return "";
    }
    return String(value).trim();
}

/**
 * Convertit une valeur en Date valide, ou null.
 * Accepte : Date, nombre (timestamp ms), chaînes ISO (AAAA-MM-JJ...),
 * JJ/MM/AAAA, JJ-MM-AAAA, JJ.MM.AAAA (heure optionnelle).
 */
export function parseDate(value: unknown): Date | null {
    if (value === null || value === undefined || value === "") {
        return null;
    }
    if (value instanceof Date) {
        return isNaN(value.getTime()) ? null : value;
    }
    if (typeof value === "number") {
        if (!isFinite(value)) {
            return null;
        }
        const d = new Date(value);
        return isNaN(d.getTime()) ? null : d;
    }
    const s = String(value).trim();
    if (!s) {
        return null;
    }
    const fr = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
    if (fr) {
        const day = +fr[1], month = +fr[2], year = +fr[3];
        const d = new Date(year, month - 1, day, fr[4] ? +fr[4] : 0, fr[5] ? +fr[5] : 0, fr[6] ? +fr[6] : 0);
        if (d.getFullYear() !== year || d.getMonth() !== month - 1 || d.getDate() !== day) {
            return null;
        }
        return d;
    }
    const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
    if (iso) {
        const year = +iso[1], month = +iso[2], day = +iso[3];
        const d = new Date(year, month - 1, day);
        if (d.getFullYear() !== year || d.getMonth() !== month - 1 || d.getDate() !== day) {
            return null;
        }
        return d;
    }
    // Formats ISO complets uniquement (pas de parsing "libre" ambigu).
    if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
        const d = new Date(s);
        return isNaN(d.getTime()) ? null : d;
    }
    return null;
}

/** Convertit en nombre fini, ou null. Accepte la virgule décimale et "45 %". */
export function parseNumber(value: unknown): number | null {
    if (value === null || value === undefined || value === "") {
        return null;
    }
    if (typeof value === "number") {
        return isFinite(value) ? value : null;
    }
    if (typeof value === "boolean" || value instanceof Date) {
        return null;
    }
    const s = String(value).trim().replace(/\s/g, "").replace("%", "").replace(",", ".");
    if (!s) {
        return null;
    }
    const n = Number(s);
    return isFinite(n) ? n : null;
}

/** Normalise une progression selon l'échelle choisie, bornée à [0, 1]. */
export function normalizeProgress(raw: number | null, scale: string): number | null {
    if (raw === null) {
        return null;
    }
    const v = scale === "0-100" ? raw / 100 : raw;
    return Math.max(0, Math.min(1, v));
}

const collator = typeof Intl !== "undefined" ? new Intl.Collator("fr", { sensitivity: "base", numeric: true }) : null;

export function compareText(a: string, b: string): number {
    return collator ? collator.compare(a, b) : (a < b ? -1 : a > b ? 1 : 0);
}

/**
 * Construit les réunions à partir des lignes brutes de la table.
 * Les lignes dont Début ou Fin est vide / invalide sont ignorées.
 * Si Fin < Début, les dates sont permutées.
 */
export function buildMeetings(rows: unknown[][], cols: RawColumns, progressScale: string): { meetings: Meeting[]; siteOrder: Map<string, number>; siteFirstSeen: Map<string, number> } {
    const meetings: Meeting[] = [];
    const siteOrder = new Map<string, number>();
    const siteFirstSeen = new Map<string, number>();
    const get = (row: unknown[], idx: number): unknown => (idx >= 0 && idx < row.length ? row[idx] : null);

    for (let i = 0; i < (rows ? rows.length : 0); i++) {
        const row = rows[i];
        if (!row) {
            continue;
        }
        let start = parseDate(get(row, cols.start));
        let end = parseDate(get(row, cols.end));
        if (!start || !end) {
            continue;
        }
        if (end.getTime() < start.getTime()) {
            const t = start; start = end; end = t;
        }
        const site = cleanText(get(row, cols.site)) || EMPTY_LABEL;
        const product = cleanText(get(row, cols.product)) || EMPTY_LABEL;
        const rawProgress = parseNumber(get(row, cols.progress));
        const order = parseNumber(get(row, cols.siteOrder));

        if (!siteFirstSeen.has(site)) {
            siteFirstSeen.set(site, siteFirstSeen.size);
        }
        if (order !== null) {
            const prev = siteOrder.get(site);
            if (prev === undefined || order < prev) {
                siteOrder.set(site, order);
            }
        }

        meetings.push({
            rowIndex: i,
            site,
            product,
            deliverable: cleanText(get(row, cols.deliverable)),
            start,
            end,
            milestone: parseDate(get(row, cols.milestone)),
            progress: normalizeProgress(rawProgress, progressScale),
            rawProgress
        });
    }
    return { meetings, siteOrder, siteFirstSeen };
}

function spanOf(list: Meeting[]): [Date | null, Date | null] {
    let min = Number.POSITIVE_INFINITY;
    let max = Number.NEGATIVE_INFINITY;
    for (const m of list) {
        min = Math.min(min, m.start.getTime());
        max = Math.max(max, m.end.getTime());
    }
    return [isFinite(min) ? new Date(min) : null, isFinite(max) ? new Date(max) : null];
}

/** Sites triés par Ordre Site croissant (sans ordre -> à la fin, puis ordre d'apparition ; jamais alphabétique). */
export function sortSites(sites: string[], siteOrder: Map<string, number>, siteFirstSeen: Map<string, number>): string[] {
    return sites.slice().sort((a, b) => {
        const oa = siteOrder.has(a) ? siteOrder.get(a) : Number.POSITIVE_INFINITY;
        const ob = siteOrder.has(b) ? siteOrder.get(b) : Number.POSITIVE_INFINITY;
        if (oa !== ob) {
            return oa < ob ? -1 : 1;
        }
        return (siteFirstSeen.get(a) ?? 0) - (siteFirstSeen.get(b) ?? 0);
    });
}

function compareMeetings(a: Meeting, b: Meeting): number {
    return (a.start.getTime() - b.start.getTime()) || (a.end.getTime() - b.end.getTime()) || compareText(a.deliverable, b.deliverable);
}

/**
 * Construit les lignes selon l'ordre des niveaux choisi, puis (mode "rows") une ligne par réunion.
 * Sites : Ordre Site croissant ; Produits : alphabétique ; Réunions : par date de début.
 * Un même produit présent sur deux sites apparaît sous chacun d'eux (et inversement).
 * `collapsed` contient les clés des groupes repliés.
 */
export function buildRows(
    meetings: Meeting[], siteOrder: Map<string, number>, siteFirstSeen: Map<string, number>,
    collapsed: Set<string>, mode: DisplayMode = "rows", order: HierarchyOrder = "site-product"
): GanttRow[] {
    const levels = order.split("-") as GroupField[];
    const rows: GanttRow[] = [];

    const walk = (list: Meeting[], depth: number, parentKey: string, parentSite: string | null): void => {
        const field = levels[depth];
        const groups = new Map<string, Meeting[]>();
        for (const m of list) {
            const v = field === "site" ? m.site : m.product;
            let g = groups.get(v);
            if (!g) {
                g = [];
                groups.set(v, g);
            }
            g.push(m);
        }
        const keys = field === "site"
            ? sortSites(Array.from(groups.keys()), siteOrder, siteFirstSeen)
            : Array.from(groups.keys()).sort(compareText);
        const leaf = depth === levels.length - 1;
        for (const value of keys) {
            const members = groups.get(value);
            const key = parentKey + "|" + (field === "site" ? "S:" : "P:") + value;
            let site = field === "site" ? value : parentSite;
            if (site === null && members.every(m => m.site === members[0].site)) {
                site = members[0].site;
            }
            const collapsible = !(leaf && mode === "compact");
            const isCollapsed = collapsible && collapsed.has(key);
            const [start, end] = spanOf(members);
            const sorted = members.slice().sort(compareMeetings);
            rows.push({ kind: "group", key, depth, field, label: value, site, meetings: sorted, leaf, collapsible, collapsed: isCollapsed, start, end });
            if (isCollapsed || (leaf && mode === "compact")) {
                continue;
            }
            if (leaf) {
                for (const m of sorted) {
                    rows.push({ kind: "meeting", key: "M|" + m.rowIndex, depth: depth + 1, meeting: m });
                }
            } else {
                walk(members, depth + 1, key, field === "site" ? value : parentSite);
            }
        }
    };
    walk(meetings, 0, "G", null);
    return rows;
}

/** Clés des groupes d'une profondeur donnée (pour "Tout replier"). */
export function groupKeysAtDepth(meetings: Meeting[], siteOrder: Map<string, number>, siteFirstSeen: Map<string, number>, mode: DisplayMode, order: HierarchyOrder, depth: number): string[] {
    return buildRows(meetings, siteOrder, siteFirstSeen, new Set<string>(), mode, order)
        .filter(r => r.kind === "group" && r.depth === depth)
        .map(r => r.key);
}

/** Plus petite et plus grande date (Début, Fin, Jalons) ; null si aucune réunion. */
export function dateExtent(meetings: Meeting[]): [Date, Date] | null {
    let min = Number.POSITIVE_INFINITY;
    let max = Number.NEGATIVE_INFINITY;
    for (const m of meetings) {
        min = Math.min(min, m.start.getTime());
        max = Math.max(max, m.end.getTime());
        if (m.milestone) {
            min = Math.min(min, m.milestone.getTime());
            max = Math.max(max, m.milestone.getTime());
        }
    }
    if (!isFinite(min) || !isFinite(max)) {
        return null;
    }
    if (max === min) {
        max = min + 24 * 3600 * 1000;
    }
    return [new Date(min), new Date(max)];
}

function pad(n: number): string {
    return n < 10 ? "0" + n : String(n);
}

/** Formate une date selon JJ-MM-AAAA (défaut), JJ/MM/AAAA ou AAAA-MM-JJ. */
export function formatDate(d: Date | null, format: string): string {
    if (!d) {
        return "";
    }
    const dd = pad(d.getDate()), mm = pad(d.getMonth() + 1), yyyy = String(d.getFullYear());
    switch (format) {
        case "DD/MM/YYYY": return `${dd}/${mm}/${yyyy}`;
        case "YYYY-MM-DD": return `${yyyy}-${mm}-${dd}`;
        default: return `${dd}-${mm}-${yyyy}`;
    }
}

/** Libellé "xx %" à partir d'une progression normalisée. */
export function formatPercent(p: number | null): string {
    if (p === null) {
        return "";
    }
    return Math.round(p * 100) + " %";
}

/**
 * Intervalle horizontal libre [a, b] pour un label, en évitant tous les jalons
 * (intervalles [l, r]). Retourne null si l'espace restant est trop petit.
 */
export function freeLabelInterval(a: number, b: number, obstacles: Array<[number, number]>, gap: number, minWidth: number): [number, number] | null {
    const sorted = obstacles.slice().sort((x, y) => x[0] - y[0]);
    for (const [l, r] of sorted) {
        if (r + gap <= a || l - gap >= b) {
            continue;
        }
        if (l - gap > a) {
            b = l - gap;
        } else {
            a = r + gap;
        }
        if (b - a < minWidth) {
            return null;
        }
    }
    return b - a >= minWidth ? [a, b] : null;
}

/** Tronque un texte avec "…" pour tenir dans maxWidth (mesure fournie). */
export function truncateText(text: string, maxWidth: number, measure: (s: string) => number): string {
    if (!text || maxWidth <= 0) {
        return "";
    }
    if (measure(text) <= maxWidth) {
        return text;
    }
    const ell = "…";
    if (measure(ell) > maxWidth) {
        return "";
    }
    let lo = 0, hi = text.length;
    while (lo < hi) {
        const mid = (lo + hi + 1) >> 1;
        if (measure(text.slice(0, mid).trimEnd() + ell) <= maxWidth) {
            lo = mid;
        } else {
            hi = mid - 1;
        }
    }
    return lo === 0 ? ell : text.slice(0, lo).trimEnd() + ell;
}

/** Couleur CSS + opacité à partir d'un hexadécimal #rgb, #rrggbb ou #rrggbbaa. */
export function splitColor(hex: string, fallbackOpacity: number): { color: string; opacity: number } {
    const s = (hex || "").trim();
    const m8 = /^#([0-9a-f]{6})([0-9a-f]{2})$/i.exec(s);
    if (m8) {
        return { color: "#" + m8[1], opacity: parseInt(m8[2], 16) / 255 };
    }
    if (/^#([0-9a-f]{6}|[0-9a-f]{3})$/i.test(s)) {
        return { color: s, opacity: fallbackOpacity };
    }
    return { color: s || "#999999", opacity: fallbackOpacity };
}
