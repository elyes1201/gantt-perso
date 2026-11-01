/*
 * Gantt PQR - niveaux de la timeline (hiérarchie de dates) et symboles de jalon.
 * Module sans accès réseau ; locale française définie localement.
 */
"use strict";

import { timeYear, timeMonth, timeMonday, timeDay, CountableTimeInterval } from "d3-time";
import { timeFormatLocale } from "d3-time-format";

/** Locale française (semaine commençant le lundi). */
export const frLocale = timeFormatLocale({
    dateTime: "%A %e %B %Y à %X",
    date: "%d/%m/%Y",
    time: "%H:%M:%S",
    periods: ["AM", "PM"],
    days: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"],
    shortDays: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."],
    months: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"],
    shortMonths: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."]
});

export type TimeUnit = "year" | "quarter" | "month" | "week" | "day";

export interface TimeLevel {
    unit: TimeUnit;
    interval: CountableTimeInterval;
    /** Libellés du plus long au plus court : on prend le premier qui tient. */
    labels: Array<(d: Date) => string>;
}

const fmtYear = frLocale.format("%Y");
const fmtMonthLong = frLocale.format("%B");
const fmtMonthShort = frLocale.format("%b");
const fmtMonthYear = frLocale.format("%b %Y");
const fmtDay = frLocale.format("%d");
const fmtDayShort = frLocale.format("%a %d");
const fmtWeek = frLocale.format("%V");
const fmtWeekYear = frLocale.format("%G");

const quarter: CountableTimeInterval = timeMonth.every(3) as CountableTimeInterval;

function quarterNum(d: Date): number {
    return Math.floor(d.getMonth() / 3) + 1;
}

export const LEVELS: Record<TimeUnit, TimeLevel> = {
    year: { unit: "year", interval: timeYear, labels: [fmtYear] },
    quarter: {
        unit: "quarter", interval: quarter,
        labels: [d => `T${quarterNum(d)} ${fmtYear(d)}`, d => `T${quarterNum(d)}`]
    },
    month: { unit: "month", interval: timeMonth, labels: [fmtMonthLong, fmtMonthShort, d => fmtMonthShort(d).charAt(0)] },
    week: {
        unit: "week", interval: timeMonday,
        labels: [d => `Semaine ${fmtWeek(d)} - ${fmtWeekYear(d)}`, d => `S${fmtWeek(d)}`, d => fmtWeek(d)]
    },
    day: { unit: "day", interval: timeDay, labels: [fmtDayShort, fmtDay] }
};

/** Libellé "mois année" pour un niveau Mois placé en haut. */
export const MONTH_TOP: TimeLevel = { unit: "month", interval: timeMonth, labels: [frLocale.format("%B %Y"), fmtMonthYear, fmtMonthShort] };

export interface HierarchyOption {
    value: string;
    displayName: string;
}

/** Hiérarchies proposées (niveau haut > niveau bas). "auto" s'adapte au zoom. */
export const HIERARCHIES: HierarchyOption[] = [
    { value: "auto", displayName: "Automatique (selon le zoom)" },
    { value: "year-quarter", displayName: "Année > Trimestre" },
    { value: "year-month", displayName: "Année > Mois" },
    { value: "quarter-month", displayName: "Trimestre > Mois" },
    { value: "month-week", displayName: "Mois > Semaine" },
    { value: "month-day", displayName: "Mois > Jour" },
    { value: "week-day", displayName: "Semaine > Jour" }
];

const DAY_MS = 24 * 3600 * 1000;

/** Choisit la hiérarchie effective ; pxPerDay = largeur d'un jour à l'écran. */
export function resolveHierarchy(value: string, pxPerDay: number): [TimeLevel, TimeLevel] {
    let v = value;
    if (!HIERARCHIES.some(h => h.value === v) || v === "auto") {
        if (pxPerDay >= 22) {
            v = "month-day";
        } else if (pxPerDay * 7 >= 45) {
            v = "month-week";
        } else if (pxPerDay * 30 >= 28) {
            v = "year-month";
        } else {
            v = "year-quarter";
        }
    }
    const [top, bottom] = v.split("-") as [TimeUnit, TimeUnit];
    return [top === "month" ? MONTH_TOP : LEVELS[top], LEVELS[bottom]];
}

export { DAY_MS };

/* --------------------------------------------------------------- jalons */

export const MILESTONE_SHAPES: HierarchyOption[] = [
    { value: "diamond", displayName: "Losange" },
    { value: "circle", displayName: "Cercle" },
    { value: "square", displayName: "Carré" },
    { value: "triangle", displayName: "Triangle" },
    { value: "star", displayName: "Étoile" },
    { value: "flag", displayName: "Drapeau" }
];

/** Chemin SVG d'un symbole de jalon centré en (0,0), de taille `size` px. */
export function milestonePath(shape: string, size: number): string {
    const h = size / 2;
    const f = (n: number) => Math.round(n * 100) / 100;
    switch (shape) {
        case "circle":
            return `M${f(-h)},0A${f(h)},${f(h)} 0 1,0 ${f(h)},0A${f(h)},${f(h)} 0 1,0 ${f(-h)},0Z`;
        case "square": {
            const s = h * 0.85;
            return `M${f(-s)},${f(-s)}H${f(s)}V${f(s)}H${f(-s)}Z`;
        }
        case "triangle":
            return `M0,${f(-h)}L${f(h)},${f(h * 0.8)}L${f(-h)},${f(h * 0.8)}Z`;
        case "star": {
            const pts: string[] = [];
            for (let i = 0; i < 10; i++) {
                const r = i % 2 === 0 ? h : h * 0.45;
                const a = -Math.PI / 2 + i * Math.PI / 5;
                pts.push(`${f(r * Math.cos(a))},${f(r * Math.sin(a))}`);
            }
            return "M" + pts.join("L") + "Z";
        }
        case "flag":
            // Mât + fanion, pied au centre de la barre
            return `M${f(-h * 0.6)},${f(h)}V${f(-h)}L${f(h)},${f(-h * 0.5)}L${f(-h * 0.6)},0`;
        default:
            return `M0,${f(-h)}L${f(h)},0L0,${f(h)}L${f(-h)},0Z`;
    }
}
