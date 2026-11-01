/*
 * Gantt PQR - panneau de paramétrage intégré au visuel (bouton ⚙).
 *
 * Les choix sont transmis au visuel via des callbacks ; le visuel les enregistre
 * dans le rapport (host.persistProperties) : aucun stockage navigateur.
 * DOM construit uniquement avec createElement / textContent (jamais innerHTML).
 */
"use strict";

import { HIERARCHIES, MILESTONE_SHAPES, milestonePath } from "./timeLevels";
import { HIERARCHY_ORDERS } from "./dataModel";

export interface Option {
    value: string;
    displayName: string;
}

export interface SiteColorState {
    site: string;
    track: string;
    fill: string;
    custom: boolean;
}

export interface PanelState {
    hierarchyOrder: string;
    showSiteBar: boolean;
    showProductBar: boolean;
    showToday: boolean;
    todayColor: string;
    todayWidth: number;
    displayMode: string;
    hierarchy: string;
    rowHeight: number;
    barHeightPercent: number;
    milestoneShape: string;
    milestoneSize: number;
    milestoneFill: string;
    labelsShow: boolean;
    labelsPosition: string;
    showProgressLabel: boolean;
    progressScale: string;
    trackOpacity: number;
    sites: SiteColorState[];
}

export type PropValue = string | number | boolean | { solid: { color: string } };

export interface PanelCallbacks {
    /** Modifie une propriété d'un objet déclaré dans capabilities.json. */
    onChange(objectName: string, propertyName: string, value: PropValue): void;
    /** Modifie la couleur (piste ou remplissage) d'un site. */
    onSiteColor(site: string, part: "track" | "fill", hex: string): void;
    /** Supprime toutes les couleurs personnalisées par site. */
    onResetSiteColors(): void;
    onClose(): void;
}

const SVG_NS = "http://www.w3.org/2000/svg";

const DISPLAY_MODES: Option[] = [
    { value: "rows", displayName: "Ligne par réunion" },
    { value: "compact", displayName: "Compact" }
];
const LABEL_POSITIONS: Option[] = [
    { value: "above", displayName: "Au-dessus" },
    { value: "insideLeft", displayName: "Intérieur gauche" }
];
const SCALES: Option[] = [
    { value: "0-1", displayName: "0 - 1" },
    { value: "0-100", displayName: "0 - 100" }
];

/** Normalise une couleur en #rrggbb pour <input type="color">. */
export function toInputColor(hex: string): string {
    const s = (hex || "").trim();
    let m = /^#([0-9a-f]{6})([0-9a-f]{2})?$/i.exec(s);
    if (m) {
        return "#" + m[1].toLowerCase();
    }
    m = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(s);
    if (m) {
        return ("#" + m[1] + m[1] + m[2] + m[2] + m[3] + m[3]).toLowerCase();
    }
    return "#999999";
}

export class SettingsPanel {
    private el: HTMLDivElement;
    private body: HTMLDivElement;
    private callbacks: PanelCallbacks;
    private openState = false;

    constructor(parent: HTMLElement, callbacks: PanelCallbacks) {
        this.callbacks = callbacks;
        this.el = document.createElement("div");
        this.el.className = "gantt-panel";
        this.el.setAttribute("role", "dialog");
        this.el.setAttribute("aria-label", "Paramètres du Gantt");
        this.el.style.display = "none";

        const header = document.createElement("div");
        header.className = "gantt-panel-header";
        const title = document.createElement("span");
        title.textContent = "Paramètres du Gantt";
        const close = document.createElement("button");
        close.type = "button";
        close.className = "gantt-btn";
        close.textContent = "×";
        close.title = "Fermer";
        close.setAttribute("aria-label", "Fermer");
        close.addEventListener("click", () => this.callbacks.onClose());
        header.appendChild(title);
        header.appendChild(close);
        this.el.appendChild(header);

        this.body = document.createElement("div");
        this.body.className = "gantt-panel-body";
        this.el.appendChild(this.body);

        // Les interactions dans le panneau ne doivent pas atteindre le Gantt.
        ["click", "mousedown", "pointerdown", "contextmenu", "mousemove"].forEach(t =>
            this.el.addEventListener(t, e => e.stopPropagation()));
        this.el.addEventListener("wheel", e => e.stopPropagation(), { passive: true });

        parent.appendChild(this.el);
    }

    public isOpen(): boolean {
        return this.openState;
    }

    public open(state: PanelState): void {
        this.openState = true;
        this.el.style.display = "flex";
        this.render(state);
    }

    public close(): void {
        this.openState = false;
        this.el.style.display = "none";
    }

    /** Rafraîchit les valeurs affichées (si ouvert), en conservant le défilement. */
    public refresh(state: PanelState): void {
        if (this.openState) {
            const top = this.body.scrollTop;
            this.render(state);
            this.body.scrollTop = top;
        }
    }

    // ------------------------------------------------------------------

    private render(s: PanelState): void {
        while (this.body.firstChild) {
            this.body.removeChild(this.body.firstChild);
        }
        const cb = this.callbacks;

        this.section("Hiérarchie");
        this.select("Niveaux", HIERARCHY_ORDERS, s.hierarchyOrder, v => cb.onChange("layout", "hierarchyOrder", v));
        this.select("Mode", DISPLAY_MODES, s.displayMode, v => cb.onChange("layout", "displayMode", v));
        this.checkbox("Barre de synthèse des Sites", s.showSiteBar, v => cb.onChange("layout", "showSiteBar", v));
        this.checkbox("Barre de synthèse des Produits", s.showProductBar, v => cb.onChange("layout", "showProductBar", v));

        this.section("Affichage");
        this.select("Hiérarchie de dates", HIERARCHIES, s.hierarchy, v => cb.onChange("timeline", "hierarchy", v));
        this.number("Hauteur de ligne (px)", s.rowHeight, 16, 80, 1, v => cb.onChange("layout", "rowHeight", v));
        this.range("Hauteur de barre (%)", s.barHeightPercent, 10, 100, 5, v => cb.onChange("layout", "barHeightPercent", v));

        this.section("Aujourd'hui");
        this.checkbox("Afficher la ligne « Aujourd'hui »", s.showToday, v => cb.onChange("timeline", "showToday", v));
        this.color("Couleur", s.todayColor, v => cb.onChange("timeline", "todayColor", { solid: { color: v } }));
        this.number("Épaisseur (px)", s.todayWidth, 1, 6, 1, v => cb.onChange("timeline", "todayWidth", v));

        this.section("Jalons");
        this.shapePicker(s.milestoneShape, s.milestoneFill, v => cb.onChange("milestones", "shape", v));
        this.number("Taille (px)", s.milestoneSize, 4, 40, 1, v => cb.onChange("milestones", "size", v));
        this.color("Couleur", s.milestoneFill, v => cb.onChange("milestones", "fill", { solid: { color: v } }));

        this.section("Labels et progression");
        this.checkbox("Afficher le nom de la réunion", s.labelsShow, v => cb.onChange("labels", "show", v));
        this.select("Position du nom", LABEL_POSITIONS, s.labelsPosition, v => cb.onChange("labels", "position", v));
        this.checkbox("Afficher « xx % »", s.showProgressLabel, v => cb.onChange("progress", "showLabel", v));
        this.select("Échelle de progression", SCALES, s.progressScale, v => cb.onChange("progress", "scale", v));

        this.section("Couleurs par site");
        const hint = document.createElement("div");
        hint.className = "gantt-panel-hint";
        hint.textContent = "Piste = barre complète (transparente) ; remplissage = progression.";
        this.body.appendChild(hint);
        const head = document.createElement("div");
        head.className = "gantt-panel-site gantt-panel-site-head";
        ["Site", "Piste", "Rempl."].forEach((t, i) => {
            const span = document.createElement("span");
            span.textContent = t;
            span.className = i === 0 ? "gantt-panel-site-name" : "gantt-panel-site-col";
            head.appendChild(span);
        });
        this.body.appendChild(head);
        for (const site of s.sites) {
            this.siteRow(site);
        }
        this.range("Opacité de la piste (%)", s.trackOpacity, 0, 100, 1, v => cb.onChange("siteColors", "trackOpacity", v));
        const reset = document.createElement("button");
        reset.type = "button";
        reset.className = "gantt-panel-reset";
        reset.textContent = "Réinitialiser les couleurs des sites";
        reset.disabled = !s.sites.some(x => x.custom);
        reset.addEventListener("click", () => cb.onResetSiteColors());
        this.body.appendChild(reset);
    }

    private section(title: string): void {
        const h = document.createElement("div");
        h.className = "gantt-panel-section";
        h.textContent = title;
        this.body.appendChild(h);
    }

    private field(label: string): HTMLLabelElement {
        const row = document.createElement("label");
        row.className = "gantt-panel-field";
        const span = document.createElement("span");
        span.textContent = label;
        row.appendChild(span);
        this.body.appendChild(row);
        return row;
    }

    private select(label: string, options: Option[], value: string, onChange: (v: string) => void): void {
        const row = this.field(label);
        const sel = document.createElement("select");
        for (const o of options) {
            const opt = document.createElement("option");
            opt.value = o.value;
            opt.textContent = o.displayName;
            sel.appendChild(opt);
        }
        sel.value = value;
        sel.addEventListener("change", () => onChange(sel.value));
        row.appendChild(sel);
    }

    private number(label: string, value: number, min: number, max: number, step: number, onChange: (v: number) => void): void {
        const row = this.field(label);
        const input = document.createElement("input");
        input.type = "number";
        input.min = String(min);
        input.max = String(max);
        input.step = String(step);
        input.value = String(value);
        input.addEventListener("change", () => {
            const v = Number(input.value);
            if (isFinite(v)) {
                onChange(Math.max(min, Math.min(max, v)));
            }
        });
        row.appendChild(input);
    }

    private range(label: string, value: number, min: number, max: number, step: number, onChange: (v: number) => void): void {
        const row = this.field(label);
        const wrap = document.createElement("span");
        wrap.className = "gantt-panel-range";
        const input = document.createElement("input");
        input.type = "range";
        input.min = String(min);
        input.max = String(max);
        input.step = String(step);
        input.value = String(value);
        const out = document.createElement("output");
        out.textContent = String(value);
        input.addEventListener("input", () => { out.textContent = input.value; });
        input.addEventListener("change", () => onChange(Number(input.value)));
        wrap.appendChild(input);
        wrap.appendChild(out);
        row.appendChild(wrap);
    }

    private checkbox(label: string, value: boolean, onChange: (v: boolean) => void): void {
        const row = this.field(label);
        row.classList.add("inline");
        const input = document.createElement("input");
        input.type = "checkbox";
        input.checked = value;
        input.addEventListener("change", () => onChange(input.checked));
        row.insertBefore(input, row.firstChild);
    }

    private color(label: string, value: string, onChange: (v: string) => void): HTMLInputElement {
        const row = this.field(label);
        const input = document.createElement("input");
        input.type = "color";
        input.value = toInputColor(value);
        input.addEventListener("change", () => onChange(input.value));
        row.appendChild(input);
        return input;
    }

    private shapePicker(value: string, fill: string, onChange: (v: string) => void): void {
        const row = this.field("Symbole");
        const group = document.createElement("span");
        group.className = "gantt-panel-shapes";
        for (const shape of MILESTONE_SHAPES) {
            const b = document.createElement("button");
            b.type = "button";
            b.className = "gantt-shape-btn" + (shape.value === value ? " selected" : "");
            b.title = shape.displayName;
            b.setAttribute("aria-label", shape.displayName);
            b.setAttribute("aria-pressed", String(shape.value === value));
            const svg = document.createElementNS(SVG_NS, "svg");
            svg.setAttribute("width", "18");
            svg.setAttribute("height", "18");
            svg.setAttribute("viewBox", "-9 -9 18 18");
            const path = document.createElementNS(SVG_NS, "path");
            path.setAttribute("d", milestonePath(shape.value, 12));
            path.setAttribute("fill", toInputColor(fill));
            path.setAttribute("stroke", "#000000");
            path.setAttribute("stroke-width", "1");
            svg.appendChild(path);
            b.appendChild(svg);
            b.addEventListener("click", (e) => {
                e.preventDefault();
                onChange(shape.value);
            });
            group.appendChild(b);
        }
        row.appendChild(group);
    }

    private siteRow(site: SiteColorState): void {
        const row = document.createElement("div");
        row.className = "gantt-panel-site";
        const name = document.createElement("span");
        name.className = "gantt-panel-site-name";
        name.textContent = site.site;
        name.title = site.site;
        row.appendChild(name);
        const mk = (part: "track" | "fill", value: string, title: string) => {
            const input = document.createElement("input");
            input.type = "color";
            input.value = toInputColor(value);
            input.title = `${site.site} - ${title}`;
            input.setAttribute("aria-label", input.title);
            input.addEventListener("change", () => this.callbacks.onSiteColor(site.site, part, input.value));
            row.appendChild(input);
        };
        mk("track", site.track, "piste");
        mk("fill", site.fill, "remplissage");
        this.body.appendChild(row);
    }
}
