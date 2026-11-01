/*
 * Gantt PQR - visuel personnalisé Power BI.
 *
 * Sécurité :
 *  - aucun accès réseau, aucun stockage navigateur, aucun log de données ;
 *  - tout texte issu des données est injecté via textContent / d3 .text()
 *    (jamais innerHTML, eval, new Function, document.write).
 */
"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { select, Selection } from "d3-selection";
import { scaleTime, ScaleTime } from "d3-scale";
import { timeYear, timeMonth, timeMonday } from "d3-time";
import { timeFormatLocale } from "d3-time-format";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;
import ITooltipService = powerbi.extensibility.ITooltipService;
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;
import DataView = powerbi.DataView;

import { VisualFormattingSettingsModel } from "./settings";
import {
    Meeting, GanttRow, ProductRow, SiteRow, RawColumns,
    buildMeetings, buildRows, dateExtent, formatDate, formatPercent,
    freeLabelInterval, truncateText, splitColor
} from "./dataModel";

type AnySelection = Selection<any, any, any, any>;

const SVG_NS = "http://www.w3.org/2000/svg";
const MARGIN_PX = 100;          // marge gauche / droite de la timeline
const HEADER_ROW_H = 24;        // hauteur de chaque niveau d'en-tête (Année, Mois)
const HEADER_H = HEADER_ROW_H * 2;
const MIN_ZOOM = 0.02;
const MAX_ZOOM = 400;
const DAY_MS = 24 * 3600 * 1000;
const DIM_OPACITY = 0.3;
const ROW_BUFFER = 10;

/** Locale française locale (aucun chargement externe). Semaine : lundi. */
const frLocale = timeFormatLocale({
    dateTime: "%A %e %B %Y à %X",
    date: "%d/%m/%Y",
    time: "%H:%M:%S",
    periods: ["AM", "PM"],
    days: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"],
    shortDays: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."],
    months: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"],
    shortMonths: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."]
});
const fmtShortMonth = frLocale.format("%b");
const fmtYear = frLocale.format("%Y");

interface SiteColor {
    track: string;
    trackOpacity: number;
    fill: string;
    fillOpacity: number;
}

export class Visual implements IVisual {
    private host: IVisualHost;
    private events: IVisualEventService;
    private selectionManager: ISelectionManager;
    private tooltipService: ITooltipService;
    private formattingSettingsService: FormattingSettingsService;
    private settings: VisualFormattingSettingsModel = new VisualFormattingSettingsModel();

    // DOM
    private root: HTMLDivElement;
    private gridPane: HTMLDivElement;
    private gridHeader: HTMLDivElement;
    private gridBody: HTMLDivElement;
    private splitter: HTMLDivElement;
    private tlPane: HTMLDivElement;
    private tlHeader: HTMLDivElement;
    private tlBody: HTMLDivElement;
    private emptyMsg: HTMLDivElement;
    private gridSvg: AnySelection;
    private gridRowsG: AnySelection;
    private headerSvg: AnySelection;
    private bodySvg: AnySelection;
    private bgRect: AnySelection;
    private gridLinesG: AnySelection;
    private rowsG: AnySelection;

    // Etat
    private viewport: powerbi.IViewport = { width: 0, height: 0 };
    private meetings: Meeting[] = [];
    private siteOrder = new Map<string, number>();
    private siteFirstSeen = new Map<string, number>();
    private rows: GanttRow[] = [];
    private collapsed = new Set<string>();
    private extent: [Date, Date] | null = null;
    private extentKey = "";
    private zoom = 1;
    private x: ScaleTime<number, number> = scaleTime();
    private contentWidth = 0;
    private gridWidthOverride: number | null = null;
    private lastSettingsGridWidth: number | null = null;
    private selectedRows = new Set<number>();
    private siteColorCache = new Map<string, SiteColor>();
    private measureCtx: CanvasRenderingContext2D | null = null;
    private renderScheduled = false;
    private syncingScroll = false;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.events = options.host.eventService;
        this.selectionManager = options.host.createSelectionManager();
        this.tooltipService = options.host.tooltipService;
        this.formattingSettingsService = new FormattingSettingsService();

        this.buildDom(options.element);

        this.selectionManager.registerOnSelectCallback(() => {
            this.refreshSelection();
        });
    }

    // ------------------------------------------------------------------ DOM

    private div(cls: string, parent: HTMLElement): HTMLDivElement {
        const d = document.createElement("div");
        d.className = cls;
        parent.appendChild(d);
        return d;
    }

    private buildDom(element: HTMLElement): void {
        this.root = this.div("gantt-root", element);

        this.gridPane = this.div("gantt-grid", this.root);
        this.gridHeader = this.div("gantt-grid-header", this.gridPane);
        const toolbar = this.div("gantt-toolbar", this.gridHeader);
        this.makeButton(toolbar, "−", "Zoom arrière", () => this.zoomBy(1 / 1.5));
        this.makeButton(toolbar, "+", "Zoom avant", () => this.zoomBy(1.5));
        this.makeButton(toolbar, "⤢", "Ajuster à la plage de dates", () => this.resetZoom());
        const colTitle = this.div("gantt-col-title", this.gridHeader);
        colTitle.textContent = "Nom";

        this.gridBody = this.div("gantt-grid-body", this.gridPane);
        this.gridSvg = select(this.gridBody).append("svg").attr("class", "gantt-grid-svg");
        this.gridRowsG = this.gridSvg.append("g").attr("class", "rows");

        this.splitter = this.div("gantt-splitter", this.root);
        this.splitter.title = "Redimensionner";
        this.initSplitter();

        this.tlPane = this.div("gantt-timeline", this.root);
        this.tlHeader = this.div("gantt-tl-header", this.tlPane);
        this.headerSvg = select(this.tlHeader).append("svg").attr("class", "gantt-header-svg");
        this.tlBody = this.div("gantt-tl-body", this.tlPane);
        this.bodySvg = select(this.tlBody).append("svg").attr("class", "gantt-body-svg");
        this.bgRect = this.bodySvg.append("rect").attr("class", "bg").attr("x", 0).attr("y", 0).attr("fill", "transparent");
        this.gridLinesG = this.bodySvg.append("g").attr("class", "gridlines");
        this.rowsG = this.bodySvg.append("g").attr("class", "rows");

        this.emptyMsg = this.div("gantt-empty", this.root);
        this.emptyMsg.textContent = "Aucune donnée valide : ajoutez Site, Produit, Début et Fin.";

        // Défilement synchronisé
        this.tlBody.addEventListener("scroll", () => {
            if (this.syncingScroll) {
                return;
            }
            this.gridBody.scrollTop = this.tlBody.scrollTop;
            this.tlHeader.scrollLeft = this.tlBody.scrollLeft;
            this.scheduleRender();
        });
        this.gridBody.addEventListener("wheel", (e: WheelEvent) => {
            e.preventDefault();
            this.tlBody.scrollTop += e.deltaY;
        }, { passive: false });

        // Zoom Ctrl + molette (ancré sous le curseur)
        const onWheel = (e: WheelEvent) => {
            if (!e.ctrlKey) {
                return;
            }
            e.preventDefault();
            const rect = this.tlBody.getBoundingClientRect();
            this.zoomBy(e.deltaY < 0 ? 1.2 : 1 / 1.2, e.clientX - rect.left);
        };
        this.tlBody.addEventListener("wheel", onWheel, { passive: false });
        this.tlHeader.addEventListener("wheel", onWheel, { passive: false });

        this.initInteractions();
    }

    private makeButton(parent: HTMLElement, label: string, title: string, onClick: () => void): void {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "gantt-btn";
        b.textContent = label;
        b.title = title;
        b.setAttribute("aria-label", title);
        b.addEventListener("click", (e) => {
            e.stopPropagation();
            onClick();
        });
        parent.appendChild(b);
    }

    private initSplitter(): void {
        let startX = 0;
        let startW = 0;
        let dragging = false;
        const move = (e: PointerEvent) => {
            if (!dragging) {
                return;
            }
            const maxW = Math.max(80, this.viewport.width - 80);
            this.gridWidthOverride = Math.max(60, Math.min(maxW, startW + e.clientX - startX));
            this.layout();
            this.renderGridRows();
        };
        const up = (e: PointerEvent) => {
            dragging = false;
            this.splitter.classList.remove("active");
            try {
                this.splitter.releasePointerCapture(e.pointerId);
            } catch {
                // capture déjà relâchée
            }
            this.render();
        };
        this.splitter.addEventListener("pointerdown", (e: PointerEvent) => {
            dragging = true;
            startX = e.clientX;
            startW = this.gridWidth();
            this.splitter.classList.add("active");
            this.splitter.setPointerCapture(e.pointerId);
            e.preventDefault();
        });
        this.splitter.addEventListener("pointermove", move);
        this.splitter.addEventListener("pointerup", up);
        this.splitter.addEventListener("pointercancel", up);
    }

    // --------------------------------------------------------------- update

    public update(options: VisualUpdateOptions): void {
        this.events.renderingStarted(options);
        try {
            this.viewport = options.viewport;
            const dataView: DataView | undefined = options.dataViews && options.dataViews[0];
            this.settings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, dataView);

            const settingsGridW = this.settings.layout.gridWidth.value;
            if (this.lastSettingsGridWidth !== settingsGridW) {
                this.gridWidthOverride = null;
                this.lastSettingsGridWidth = settingsGridW;
            }
            this.siteColorCache.clear();

            this.readData(dataView);
            this.refreshSelectionSet();
            this.render();
            this.events.renderingFinished(options);
        } catch (err) {
            // Aucune donnée n'est journalisée.
            this.meetings = [];
            this.rows = [];
            this.showEmpty(true);
            this.events.renderingFailed(options);
        }
    }

    private readData(dataView: DataView | undefined): void {
        const table = dataView && dataView.table;
        this.meetings = [];
        if (!table || !table.columns || !table.rows) {
            this.rows = [];
            this.extent = null;
            return;
        }
        const cols: RawColumns = { site: -1, product: -1, deliverable: -1, start: -1, end: -1, milestone: -1, progress: -1, siteOrder: -1 };
        table.columns.forEach((c, i) => {
            const roles = c.roles || {};
            (Object.keys(cols) as Array<keyof RawColumns>).forEach(k => {
                if (roles[k] && cols[k] < 0) {
                    cols[k] = i;
                }
            });
        });
        const scale = String(this.settings.progress.scale.value?.value ?? "0-1");
        const res = buildMeetings(table.rows as unknown[][], cols, scale);
        for (const m of res.meetings) {
            m.selectionId = this.host.createSelectionIdBuilder().withTable(table, m.rowIndex).createSelectionId();
        }
        this.meetings = res.meetings;
        this.siteOrder = res.siteOrder;
        this.siteFirstSeen = res.siteFirstSeen;
        this.rows = buildRows(this.meetings, this.siteOrder, this.siteFirstSeen, this.collapsed);

        const ext = dateExtent(this.meetings);
        const key = ext ? ext[0].getTime() + "|" + ext[1].getTime() : "";
        if (key !== this.extentKey) {
            this.zoom = 1;
            this.extentKey = key;
        }
        this.extent = ext;
    }

    // ------------------------------------------------------------- réglages

    private rowHeight(): number {
        return Math.max(16, this.settings.layout.rowHeight.value || 30);
    }

    private barHeight(): number {
        const pct = Math.max(10, Math.min(100, this.settings.layout.barHeightPercent.value || 60));
        return Math.max(2, Math.round(this.rowHeight() * pct / 100));
    }

    private gridWidth(): number {
        const w = this.gridWidthOverride ?? this.settings.layout.gridWidth.value ?? 200;
        return Math.max(60, Math.min(w, Math.max(60, this.viewport.width - 80)));
    }

    private milestoneSize(): number {
        return Math.max(4, this.settings.milestones.size.value || 10);
    }

    private labelFontSize(): number {
        return Math.max(6, this.settings.labels.font.fontSize.value || 10);
    }

    private labelFontFamily(): string {
        return this.settings.labels.font.fontFamily.value || "Segoe UI, sans-serif";
    }

    private gridFontSize(): number {
        return Math.max(6, this.settings.grid.font.fontSize.value || 11);
    }

    private gridFontFamily(): string {
        return this.settings.grid.font.fontFamily.value || "Segoe UI, sans-serif";
    }

    private dateFormat(): string {
        return String(this.settings.timeline.dateFormat.value?.value ?? "DD-MM-YYYY");
    }

    private colorOf(site: string): SiteColor {
        const cached = this.siteColorCache.get(site);
        if (cached) {
            return cached;
        }
        const sc = this.settings.siteColors;
        const trackFallback = Math.max(0, Math.min(100, sc.trackOpacity.value ?? 33)) / 100;
        const norm = site.trim().toLowerCase();
        let trackHex: string;
        let fillHex: string;
        if (norm === "lille") {
            trackHex = sc.lilleTrack.value.value;
            fillHex = sc.lilleFill.value.value;
        } else if (norm === "ulis" || norm === "les ulis") {
            trackHex = sc.ulisTrack.value.value;
            fillHex = sc.ulisFill.value.value;
        } else {
            // Index stable : rang du site parmi les "autres sites" dans l'ordre d'affichage.
            const others = this.rows.filter(r => r.kind === "site" && !["lille", "ulis", "les ulis"].includes(r.site.trim().toLowerCase())).map(r => r.site);
            const idx = Math.max(0, others.indexOf(site)) % sc.palette.length;
            trackHex = sc.palette[idx][0].value.value;
            fillHex = sc.palette[idx][1].value.value;
        }
        const t = splitColor(trackHex, trackFallback);
        const f = splitColor(fillHex, 1);
        const res: SiteColor = { track: t.color, trackOpacity: t.opacity, fill: f.color, fillOpacity: f.opacity };
        this.siteColorCache.set(site, res);
        return res;
    }

    private measure(text: string, fontSize: number, fontFamily: string): number {
        if (!this.measureCtx) {
            const canvas = document.createElement("canvas");
            this.measureCtx = canvas.getContext("2d");
        }
        if (!this.measureCtx) {
            return text.length * fontSize * 0.55;
        }
        this.measureCtx.font = `${fontSize}px ${fontFamily}`;
        return this.measureCtx.measureText(text).width;
    }

    // ----------------------------------------------------------------- zoom

    private timelineViewportWidth(): number {
        return Math.max(50, this.viewport.width - this.gridWidth() - 5);
    }

    private basePxPerMs(): number {
        if (!this.extent) {
            return 1 / DAY_MS;
        }
        const span = Math.max(DAY_MS, this.extent[1].getTime() - this.extent[0].getTime());
        const usable = Math.max(50, this.timelineViewportWidth() - 2 * MARGIN_PX);
        return usable / span;
    }

    private updateScale(): void {
        if (!this.extent) {
            this.contentWidth = this.timelineViewportWidth();
            return;
        }
        const span = Math.max(DAY_MS, this.extent[1].getTime() - this.extent[0].getTime());
        const inner = span * this.basePxPerMs() * this.zoom;
        this.x = scaleTime().domain([this.extent[0], new Date(this.extent[0].getTime() + span)]).range([MARGIN_PX, MARGIN_PX + inner]);
        this.contentWidth = Math.ceil(inner + 2 * MARGIN_PX);
    }

    private zoomBy(factor: number, anchorX?: number): void {
        if (!this.extent) {
            return;
        }
        const ox = anchorX ?? this.tlBody.clientWidth / 2;
        const anchorDate = this.x.invert(this.tlBody.scrollLeft + ox);
        const nz = Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, this.zoom * factor));
        // Plafond : 400 px par jour
        const maxZoom = (400 / DAY_MS) / this.basePxPerMs();
        this.zoom = Math.min(nz, Math.max(1, maxZoom));
        this.updateScale();
        this.render(false);
        this.tlBody.scrollLeft = Math.max(0, this.x(anchorDate) - ox);
        this.tlHeader.scrollLeft = this.tlBody.scrollLeft;
        this.scheduleRender();
    }

    private resetZoom(): void {
        this.zoom = 1;
        this.render();
        this.tlBody.scrollLeft = 0;
        this.tlHeader.scrollLeft = 0;
    }

    // --------------------------------------------------------------- rendu

    private showEmpty(empty: boolean): void {
        this.emptyMsg.style.display = empty ? "flex" : "none";
        this.gridPane.style.display = empty ? "none" : "block";
        this.splitter.style.display = empty ? "none" : "block";
        this.tlPane.style.display = empty ? "none" : "block";
    }

    private layout(): void {
        const w = this.viewport.width;
        const h = this.viewport.height;
        const gw = this.gridWidth();
        this.root.style.width = w + "px";
        this.root.style.height = h + "px";
        this.root.style.fontFamily = this.gridFontFamily();

        this.gridPane.style.width = gw + "px";
        this.gridPane.style.height = h + "px";
        this.gridHeader.style.height = HEADER_H + "px";
        this.gridBody.style.top = HEADER_H + "px";
        this.gridBody.style.height = Math.max(0, h - HEADER_H) + "px";

        this.splitter.style.left = gw + "px";
        this.splitter.style.height = h + "px";

        const tlLeft = gw + 5;
        this.tlPane.style.left = tlLeft + "px";
        this.tlPane.style.width = Math.max(0, w - tlLeft) + "px";
        this.tlPane.style.height = h + "px";
        this.tlHeader.style.height = HEADER_H + "px";
        this.tlBody.style.top = HEADER_H + "px";
        this.tlBody.style.height = Math.max(0, h - HEADER_H) + "px";

        const tl = this.settings.timeline;
        this.gridHeader.style.background = tl.yearBackground.value.value;
        this.gridHeader.style.color = tl.yearTextColor.value.value;
    }

    private render(scrollRender = true): void {
        const empty = this.meetings.length === 0 || !this.extent;
        this.showEmpty(empty);
        if (empty) {
            this.gridRowsG.selectAll("*").remove();
            this.rowsG.selectAll("*").remove();
            this.headerSvg.selectAll("*").remove();
            this.gridLinesG.selectAll("*").remove();
            return;
        }
        this.layout();
        this.updateScale();

        const rh = this.rowHeight();
        const totalH = this.rows.length * rh;
        const bodyH = Math.max(totalH, 1);
        this.bodySvg.attr("width", this.contentWidth).attr("height", bodyH);
        this.bgRect.attr("width", this.contentWidth).attr("height", Math.max(bodyH, this.tlBody.clientHeight));

        // Compense la barre de défilement horizontale de la timeline.
        const hScrollbar = Math.max(0, this.tlBody.offsetHeight - this.tlBody.clientHeight);
        this.gridSvg.attr("width", this.gridWidth()).attr("height", bodyH + hScrollbar);

        this.renderHeader();
        this.renderGridLines(bodyH);
        if (scrollRender) {
            this.renderVisibleRows();
        }
    }

    private scheduleRender(): void {
        if (this.renderScheduled) {
            return;
        }
        this.renderScheduled = true;
        requestAnimationFrame(() => {
            this.renderScheduled = false;
            this.renderVisibleRows();
        });
    }

    private visibleRange(): [number, number] {
        const rh = this.rowHeight();
        const top = this.tlBody.scrollTop;
        const h = this.tlBody.clientHeight || this.viewport.height;
        const first = Math.max(0, Math.floor(top / rh) - ROW_BUFFER);
        const last = Math.min(this.rows.length - 1, Math.ceil((top + h) / rh) + ROW_BUFFER);
        return [first, last];
    }

    private renderVisibleRows(): void {
        this.renderGridRows();
        this.renderTimelineRows();
    }

    // ------------------------------------------------------- en-tête dates

    private renderHeader(): void {
        const tl = this.settings.timeline;
        const w = this.contentWidth;
        const svg = this.headerSvg.attr("width", w).attr("height", HEADER_H);
        svg.selectAll("*").remove();

        const d0 = this.x.invert(0);
        const d1 = this.x.invert(w);
        const fontFamily = this.gridFontFamily();

        svg.append("rect").attr("x", 0).attr("y", 0).attr("width", w).attr("height", HEADER_ROW_H).attr("fill", tl.yearBackground.value.value);
        svg.append("rect").attr("x", 0).attr("y", HEADER_ROW_H).attr("width", w).attr("height", HEADER_ROW_H).attr("fill", tl.monthBackground.value.value);

        // Années
        const years = timeYear.range(timeYear.floor(d0), timeYear.offset(timeYear.floor(d1), 1));
        const yg = svg.append("g").attr("class", "years");
        years.forEach(y => {
            const xs = Math.max(0, this.x(y));
            const xe = Math.min(w, this.x(timeYear.offset(y, 1)));
            if (xe <= xs) {
                return;
            }
            yg.append("line").attr("x1", this.x(y)).attr("x2", this.x(y)).attr("y1", 0).attr("y2", HEADER_ROW_H).attr("stroke", tl.gridLineColor.value.value);
            const label = fmtYear(y);
            const avail = xe - xs - 8;
            const txt = truncateText(label, avail, s => this.measure(s, 11, fontFamily));
            if (txt) {
                yg.append("text").attr("x", xs + 4).attr("y", HEADER_ROW_H / 2).attr("dy", "0.35em")
                    .attr("fill", tl.yearTextColor.value.value).style("font-size", "11px").style("font-weight", "600").text(txt);
            }
        });

        // Mois (libellés courts français)
        const months = timeMonth.range(timeMonth.floor(d0), timeMonth.offset(timeMonth.floor(d1), 1));
        const mg = svg.append("g").attr("class", "months");
        months.forEach(m => {
            const xs = this.x(m);
            const xe = this.x(timeMonth.offset(m, 1));
            if (xe < 0 || xs > w) {
                return;
            }
            mg.append("line").attr("x1", xs).attr("x2", xs).attr("y1", HEADER_ROW_H).attr("y2", HEADER_H).attr("stroke", tl.gridLineColor.value.value);
            const txt = truncateText(fmtShortMonth(m), xe - xs - 6, s => this.measure(s, 10, fontFamily));
            if (txt) {
                mg.append("text").attr("x", xs + 3).attr("y", HEADER_ROW_H * 1.5).attr("dy", "0.35em")
                    .attr("fill", tl.monthTextColor.value.value).style("font-size", "10px").text(txt);
            }
        });
        svg.append("line").attr("x1", 0).attr("x2", w).attr("y1", HEADER_H - 0.5).attr("y2", HEADER_H - 0.5).attr("stroke", "#d0d0d0");
        svg.append("line").attr("x1", 0).attr("x2", w).attr("y1", HEADER_ROW_H - 0.5).attr("y2", HEADER_ROW_H - 0.5).attr("stroke", "#e0e0e0");
    }

    private renderGridLines(height: number): void {
        const color = this.settings.timeline.gridLineColor.value.value;
        const d0 = this.x.invert(0);
        const d1 = this.x.invert(this.contentWidth);
        const weekPx = this.x(new Date(d0.getTime() + 7 * DAY_MS)) - this.x(d0);
        // Semaines (début lundi) si assez de place, sinon mois.
        const ticks = weekPx >= 25 ? timeMonday.range(timeMonday.floor(d0), d1) : timeMonth.range(timeMonth.floor(d0), d1);
        const sel = this.gridLinesG.selectAll<SVGLineElement, Date>("line").data(ticks, (d: Date) => String(d.getTime()));
        sel.exit().remove();
        sel.enter().append("line")
            .merge(sel)
            .attr("x1", d => Math.round(this.x(d)) + 0.5)
            .attr("x2", d => Math.round(this.x(d)) + 0.5)
            .attr("y1", 0)
            .attr("y2", Math.max(height, this.tlBody.clientHeight))
            .attr("stroke", color)
            .attr("stroke-width", 1)
            .attr("stroke-dasharray", "2,2");
    }

    // --------------------------------------------------------- grille gauche

    private renderGridRows(): void {
        const rh = this.rowHeight();
        const gw = this.gridWidth();
        const [first, last] = this.visibleRange();
        const visible = this.rows.slice(first, last + 1).map((r, i) => ({ row: r, index: first + i }));
        const fontSize = this.gridFontSize();
        const fontFamily = this.gridFontFamily();
        const textColor = this.settings.grid.textColor.value.value;
        const siteBg = this.settings.grid.siteBackground.value.value;

        this.gridSvg.attr("width", gw);
        const sel = this.gridRowsG.selectAll<SVGGElement, { row: GanttRow; index: number }>("g.grow")
            .data(visible, d => d.row.key);
        sel.exit().remove();
        const enter = sel.enter().append("g").attr("class", "grow");
        enter.append("rect").attr("class", "grow-bg");
        enter.append("line").attr("class", "grow-sep");
        enter.append("text").attr("class", "grow-text");
        enter.append("title");

        const all = enter.merge(sel);
        all.attr("transform", d => `translate(0,${d.index * rh})`)
            .classed("site", d => d.row.kind === "site")
            .classed("product", d => d.row.kind === "product")
            .style("cursor", "pointer")
            .style("opacity", d => this.rowOpacity(d.row));
        all.select("rect.grow-bg")
            .attr("width", gw).attr("height", rh)
            .attr("fill", d => d.row.kind === "site" ? siteBg : "transparent");
        all.select("line.grow-sep")
            .attr("x1", 0).attr("x2", gw).attr("y1", rh - 0.5).attr("y2", rh - 0.5).attr("stroke", "#ececec");
        all.select("text.grow-text")
            .attr("x", d => d.row.kind === "site" ? 6 : 22)
            .attr("y", rh / 2).attr("dy", "0.35em")
            .attr("fill", textColor)
            .style("font-size", fontSize + "px")
            .style("font-family", fontFamily)
            .style("font-weight", d => d.row.kind === "site" ? "600" : "400")
            .text(d => {
                const label = d.row.kind === "site"
                    ? ((d.row as SiteRow).collapsed ? "▸ " : "▾ ") + d.row.site
                    : (d.row as ProductRow).product;
                const avail = gw - (d.row.kind === "site" ? 6 : 22) - 6;
                return truncateText(label, avail, s => this.measure(s, fontSize, fontFamily));
            });
        all.select("title").text(d => d.row.kind === "site" ? d.row.site : (d.row as ProductRow).product);
    }

    // -------------------------------------------------------------- timeline

    private renderTimelineRows(): void {
        const rh = this.rowHeight();
        const w = this.contentWidth;
        const [first, last] = this.visibleRange();
        const visible = this.rows.slice(first, last + 1).map((r, i) => ({ row: r, index: first + i }));
        const siteBg = this.settings.grid.siteBackground.value.value;

        const sel = this.rowsG.selectAll<SVGGElement, { row: GanttRow; index: number }>("g.trow")
            .data(visible, d => d.row.key);
        sel.exit().remove();
        const enter = sel.enter().append("g").attr("class", "trow");
        enter.append("rect").attr("class", "trow-bg");
        enter.append("line").attr("class", "trow-sep");
        enter.append("g").attr("class", "layer-tracks");
        enter.append("g").attr("class", "layer-lines");
        enter.append("g").attr("class", "layer-labels");
        enter.append("g").attr("class", "layer-milestones");

        const all = enter.merge(sel);
        all.attr("transform", d => `translate(0,${d.index * rh})`);
        all.select("rect.trow-bg").attr("width", w).attr("height", rh)
            .attr("fill", d => d.row.kind === "site" ? siteBg : "transparent")
            .attr("fill-opacity", d => d.row.kind === "site" ? 0.6 : 0);
        all.select("line.trow-sep").attr("x1", 0).attr("x2", w).attr("y1", rh - 0.5).attr("y2", rh - 0.5).attr("stroke", "#f2f2f2");

        const self = this;
        all.each(function (d) {
            self.renderRowContent(select(this), d.row);
        });
    }

    private renderRowContent(g: AnySelection, row: GanttRow): void {
        const rh = this.rowHeight();
        const bh = this.barHeight();
        const ms = this.milestoneSize();
        const labelsOn = this.settings.labels.show.value;
        const above = String(this.settings.labels.position.value?.value) === "above";
        const fontSize = this.labelFontSize();
        const fontFamily = this.labelFontFamily();
        const labelColor = this.settings.labels.color.value.value;
        const showProgress = this.settings.progress.showLabel.value;
        const msFill = this.settings.milestones.fill.value.value;
        const msStroke = this.settings.milestones.stroke.value.value;
        const lineColor = this.settings.milestones.lineColor.value.value;

        const tracks = g.select("g.layer-tracks");
        const lines = g.select("g.layer-lines");
        const labels = g.select("g.layer-labels");
        const milestones = g.select("g.layer-milestones");

        if (row.kind === "site") {
            // Barre de synthèse fine (étendue du site)
            const data = row.start && row.end ? [row] : [];
            const c = this.colorOf(row.site);
            const s = tracks.selectAll<SVGRectElement, SiteRow>("rect.summary").data(data);
            s.exit().remove();
            s.enter().append("rect").attr("class", "summary").merge(s)
                .attr("x", d => this.x(d.start))
                .attr("width", d => Math.max(2, this.x(d.end) - this.x(d.start)))
                .attr("y", rh / 2 - 2).attr("height", 4).attr("rx", 2)
                .attr("fill", c.fill).attr("fill-opacity", 0.45);
            lines.selectAll("*").remove();
            labels.selectAll("*").remove();
            milestones.selectAll("*").remove();
            return;
        }

        // Position verticale de la barre
        const barTop = above ? Math.max(0, rh - bh - 2) : Math.round((rh - bh) / 2);
        const cy = barTop + bh / 2;
        const meetings = row.meetings;
        const geo = meetings.map(m => {
            const x0 = this.x(m.start);
            const x1 = Math.max(x0 + 2, this.x(m.end));
            const xm = m.milestone ? this.x(m.milestone) : null;
            return { m, x0, x1, xm };
        });
        const obstacles: Array<[number, number]> = geo.filter(d => d.xm !== null).map(d => [d.xm - ms / 2 - 1, d.xm + ms / 2 + 1] as [number, number]);

        // Pistes + remplissage
        const tg = tracks.selectAll<SVGGElement, typeof geo[0]>("g.meeting").data(geo, d => String(d.m.rowIndex));
        tg.exit().remove();
        const tgEnter = tg.enter().append("g").attr("class", "meeting");
        tgEnter.append("rect").attr("class", "track");
        tgEnter.append("rect").attr("class", "fill");
        const tgAll = tgEnter.merge(tg);
        tgAll.style("opacity", d => this.meetingOpacity(d.m));
        tgAll.select("rect.track")
            .attr("x", d => d.x0).attr("y", barTop).attr("width", d => d.x1 - d.x0).attr("height", bh).attr("rx", 2)
            .attr("fill", d => this.colorOf(d.m.site).track)
            .attr("fill-opacity", d => this.colorOf(d.m.site).trackOpacity);
        tgAll.select("rect.fill")
            .attr("x", d => d.x0).attr("y", barTop).attr("height", bh).attr("rx", 2)
            .attr("width", d => (d.x1 - d.x0) * (d.m.progress ?? 0))
            .attr("fill", d => this.colorOf(d.m.site).fill)
            .attr("fill-opacity", d => this.colorOf(d.m.site).fillOpacity);

        // Traits pointillés Fin -> Jalon (uniquement si Jalon après Fin)
        const lineData = geo.filter(d => d.xm !== null && d.m.milestone.getTime() > d.m.end.getTime());
        const ls = lines.selectAll<SVGLineElement, typeof geo[0]>("line.ms-line").data(lineData, d => String(d.m.rowIndex));
        ls.exit().remove();
        ls.enter().append("line").attr("class", "ms-line").merge(ls)
            .attr("x1", d => d.x1).attr("x2", d => d.xm)
            .attr("y1", cy).attr("y2", cy)
            .attr("stroke", lineColor).attr("stroke-width", 1).attr("stroke-dasharray", "4,4")
            .style("opacity", d => this.meetingOpacity(d.m));

        // Labels : Livrable CPV + "xx %"
        interface LabelItem { key: string; m: Meeting; x: number; y: number; text: string; cls: string; anchorBaseline: string; bg: boolean; w: number }
        const labelItems: LabelItem[] = [];
        const measure = (s: string) => this.measure(s, fontSize, fontFamily);
        const sortedGeo = geo.slice().sort((a, b) => a.x0 - b.x0);
        const sortedObstacles = obstacles.slice().sort((p, q) => p[0] - q[0]);
        // 1) "xx %" : à droite de la barre, sans recouvrir un jalon ni la barre suivante ;
        //    à défaut, à l'intérieur droit de la barre ; sinon masqué.
        const progressBoxes: Array<[number, number]> = [];
        sortedGeo.forEach((d, i) => {
            if (!showProgress || d.m.progress === null) {
                return;
            }
            const text = formatPercent(d.m.progress);
            const tw = measure(text);
            const next = sortedGeo[i + 1];
            const limit = next ? next.x0 - 2 : Number.POSITIVE_INFINITY;
            let px = d.x1 + 4;
            for (const [l, r] of sortedObstacles) {
                if (px < r + 1 && px + tw > l - 1) {
                    px = r + 3;
                }
            }
            if (px + tw > limit) {
                // Repli : intérieur droit de la barre
                px = d.x1 - tw - 3;
                const hit = sortedObstacles.some(([l, r]) => px < r + 1 && px + tw > l - 1);
                if (px < d.x0 + 2 || hit) {
                    return;
                }
            }
            progressBoxes.push([px - 1, px + tw + 1]);
            labelItems.push({ key: "p" + d.m.rowIndex, m: d.m, x: px, y: cy, text, cls: "lbl-progress", anchorBaseline: "central", bg: true, w: tw });
        });

        // 2) Livrable CPV : au-dessus ou intérieur gauche, tronqué, jamais sur un jalon.
        sortedGeo.forEach((d, i) => {
            if (!labelsOn || !d.m.deliverable) {
                return;
            }
            let a: number, b: number;
            if (above) {
                const next = sortedGeo[i + 1];
                a = d.x0;
                b = next ? Math.max(d.x1, next.x0 - 4) : this.contentWidth - 2;
            } else {
                a = d.x0 + 4;
                b = d.x1 - 4;
            }
            const avoid = above ? obstacles : obstacles.concat(progressBoxes);
            const free = freeLabelInterval(a, b, avoid, 2, 8);
            if (!free) {
                return;
            }
            const text = truncateText(d.m.deliverable, free[1] - free[0], measure);
            if (text) {
                labelItems.push({
                    key: "d" + d.m.rowIndex, m: d.m, x: free[0],
                    y: above ? barTop - 2 : cy, text, cls: "lbl-deliverable",
                    anchorBaseline: above ? "auto" : "central", bg: false, w: measure(text)
                });
            }
        });

        const lbl = labels.selectAll<SVGGElement, LabelItem>("g.lbl").data(labelItems, d => d.key);
        lbl.exit().remove();
        const lblEnter = lbl.enter().append("g").attr("class", "lbl");
        lblEnter.append("rect").attr("class", "lbl-bg");
        lblEnter.append("text");
        const lblAll = lblEnter.merge(lbl);
        lblAll.style("opacity", d => this.meetingOpacity(d.m));
        lblAll.select("rect.lbl-bg")
            .attr("display", d => d.bg ? null : "none")
            .attr("x", d => d.x - 1).attr("y", cy - fontSize / 2 - 1)
            .attr("width", d => d.w + 2).attr("height", fontSize + 2)
            .attr("fill", "#ffffff").attr("fill-opacity", 0.85);
        lblAll.select("text")
            .attr("class", d => d.cls)
            .attr("x", d => d.x).attr("y", d => d.y)
            .attr("dominant-baseline", d => d.anchorBaseline)
            .attr("fill", labelColor)
            .style("font-size", fontSize + "px")
            .style("font-family", fontFamily)
            .text(d => d.text);

        // Jalons (dessinés en dernier : au-dessus des labels)
        const msData = geo.filter(d => d.xm !== null);
        const h = ms / 2;
        const mss = milestones.selectAll<SVGPathElement, typeof geo[0]>("path.milestone").data(msData, d => String(d.m.rowIndex));
        mss.exit().remove();
        mss.enter().append("path").attr("class", "milestone").merge(mss)
            .attr("d", `M0,${-h}L${h},0L0,${h}L${-h},0Z`)
            .attr("transform", d => `translate(${d.xm},${cy})`)
            .attr("fill", msFill).attr("stroke", msStroke).attr("stroke-width", 1)
            .style("opacity", d => this.meetingOpacity(d.m));
    }

    // ---------------------------------------------------------- sélection

    private refreshSelectionSet(): void {
        this.selectedRows.clear();
        const ids = this.selectionManager.getSelectionIds() as ISelectionId[];
        if (!ids || ids.length === 0) {
            return;
        }
        for (const m of this.meetings) {
            if (ids.some(id => id.equals(m.selectionId))) {
                this.selectedRows.add(m.rowIndex);
            }
        }
    }

    private refreshSelection(): void {
        this.refreshSelectionSet();
        this.renderVisibleRows();
    }

    private meetingOpacity(m: Meeting): number {
        if (this.selectedRows.size === 0) {
            return 1;
        }
        return this.selectedRows.has(m.rowIndex) ? 1 : DIM_OPACITY;
    }

    private rowOpacity(row: GanttRow): number {
        if (this.selectedRows.size === 0) {
            return 1;
        }
        if (row.kind === "site") {
            return 1;
        }
        return row.meetings.some(m => this.selectedRows.has(m.rowIndex)) ? 1 : DIM_OPACITY;
    }

    private meetingFromEvent(e: Event): Meeting | null {
        let node = e.target as Element | null;
        while (node && node !== this.tlBody) {
            const d = (select(node).datum() as any);
            if (d && d.m && d.m.selectionId) {
                return d.m as Meeting;
            }
            node = node.parentElement;
        }
        return null;
    }

    private rowFromGridEvent(e: Event): GanttRow | null {
        let node = e.target as Element | null;
        while (node && node !== this.gridBody) {
            const d = (select(node).datum() as any);
            if (d && d.row && d.row.key) {
                return d.row as GanttRow;
            }
            node = node.parentElement;
        }
        return null;
    }

    private applySelection(ids: ISelectionId[], multi: boolean): void {
        if (!this.host.hostCapabilities || this.host.hostCapabilities.allowInteractions !== false) {
            this.selectionManager.select(ids, multi).then(() => this.refreshSelection());
        }
    }

    private initInteractions(): void {
        // Timeline : clic, infobulles, menu contextuel
        this.tlBody.addEventListener("click", (e: MouseEvent) => {
            const m = this.meetingFromEvent(e);
            if (m) {
                this.applySelection([m.selectionId], e.ctrlKey || e.metaKey);
            } else {
                this.selectionManager.clear().then(() => this.refreshSelection());
            }
            e.stopPropagation();
        });
        this.tlBody.addEventListener("contextmenu", (e: MouseEvent) => {
            const m = this.meetingFromEvent(e);
            this.selectionManager.showContextMenu(m ? m.selectionId : ({} as ISelectionId), { x: e.clientX, y: e.clientY });
            e.preventDefault();
        });
        this.tlBody.addEventListener("mousemove", (e: MouseEvent) => {
            if (!this.tooltipService || !this.tooltipService.enabled()) {
                return;
            }
            const m = this.meetingFromEvent(e);
            if (!m) {
                this.tooltipService.hide({ isTouchEvent: false, immediately: true });
                return;
            }
            this.tooltipService.show({
                dataItems: this.tooltipItems(m),
                identities: [m.selectionId],
                coordinates: [e.clientX, e.clientY],
                isTouchEvent: false
            });
        });
        this.tlBody.addEventListener("mouseleave", () => {
            if (this.tooltipService) {
                this.tooltipService.hide({ isTouchEvent: false, immediately: true });
            }
        });

        // Grille : repli / dépli des sites, sélection des produits
        this.gridBody.addEventListener("click", (e: MouseEvent) => {
            const row = this.rowFromGridEvent(e);
            if (!row) {
                this.selectionManager.clear().then(() => this.refreshSelection());
                return;
            }
            if (row.kind === "site") {
                if (this.collapsed.has(row.site)) {
                    this.collapsed.delete(row.site);
                } else {
                    this.collapsed.add(row.site);
                }
                this.rows = buildRows(this.meetings, this.siteOrder, this.siteFirstSeen, this.collapsed);
                this.render();
            } else {
                this.applySelection(row.meetings.map(m => m.selectionId), e.ctrlKey || e.metaKey);
            }
        });
        this.gridBody.addEventListener("contextmenu", (e: MouseEvent) => {
            const row = this.rowFromGridEvent(e);
            const id = row && row.kind === "product" && row.meetings.length ? row.meetings[0].selectionId : ({} as ISelectionId);
            this.selectionManager.showContextMenu(id, { x: e.clientX, y: e.clientY });
            e.preventDefault();
        });
    }

    private tooltipItems(m: Meeting): VisualTooltipDataItem[] {
        const fmt = this.dateFormat();
        return [
            { displayName: "Site", value: m.site },
            { displayName: "Produit", value: m.product },
            { displayName: "Livrable CPV", value: m.deliverable },
            { displayName: "Début", value: formatDate(m.start, fmt) },
            { displayName: "Fin", value: formatDate(m.end, fmt) },
            { displayName: "Jalons", value: formatDate(m.milestone, fmt) },
            { displayName: "Progress", value: formatPercent(m.progress) }
        ];
    }

    // ---------------------------------------------------------- format pane

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.settings);
    }
}
