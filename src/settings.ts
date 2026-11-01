/*
 * Gantt PQR - modèle du volet Format (formattingModel).
 * Chaque carte correspond à un objet déclaré dans capabilities.json.
 */
"use strict";

import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import { HIERARCHIES, MILESTONE_SHAPES } from "./timeLevels";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

function color(name: string, displayName: string, value: string): formattingSettings.ColorPicker {
    return new formattingSettings.ColorPicker({ name, displayName, value: { value } });
}

function num(name: string, displayName: string, value: number, min: number, max: number): formattingSettings.NumUpDown {
    return new formattingSettings.NumUpDown({
        name,
        displayName,
        value,
        options: {
            minValue: { type: powerbi.visuals.ValidatorType.Min, value: min },
            maxValue: { type: powerbi.visuals.ValidatorType.Max, value: max }
        }
    });
}

export const DISPLAY_MODES: powerbi.IEnumMember[] = [
    { value: "rows", displayName: "Une ligne par réunion (Site > Produit > Réunion)" },
    { value: "compact", displayName: "Compact (réunions sur la ligne Produit)" }
];

/** Disposition : mode d'affichage, hauteur de ligne, hauteur de barre, largeur de la grille. */
export class LayoutCard extends FormattingSettingsCard {
    displayMode = new formattingSettings.ItemDropdown({ name: "displayMode", displayName: "Mode d'affichage", items: DISPLAY_MODES, value: DISPLAY_MODES[0] });
    rowHeight = num("rowHeight", "Hauteur de ligne (px)", 30, 16, 80);
    barHeightPercent = num("barHeightPercent", "Hauteur de barre (% de la ligne)", 60, 10, 100);
    gridWidth = num("gridWidth", "Largeur de la grille (px)", 200, 60, 800);
    showSettingsButton = new formattingSettings.ToggleSwitch({ name: "showSettingsButton", displayName: "Bouton de paramétrage (⚙) dans le visuel", value: true });

    name = "layout";
    displayName = "Disposition";
    slices: FormattingSettingsSlice[] = [this.displayMode, this.rowHeight, this.barHeightPercent, this.gridWidth, this.showSettingsButton];
}

export const SHAPE_ITEMS: powerbi.IEnumMember[] = MILESTONE_SHAPES.map(s => ({ value: s.value, displayName: s.displayName }));

/** Jalons : symbole, taille, couleurs. */
export class MilestoneCard extends FormattingSettingsCard {
    shape = new formattingSettings.ItemDropdown({ name: "shape", displayName: "Symbole du jalon", items: SHAPE_ITEMS, value: SHAPE_ITEMS[0] });
    size = num("size", "Taille du jalon (px)", 10, 4, 40);
    fill = color("fill", "Couleur du jalon", "#f44a53");
    stroke = color("stroke", "Bordure du jalon", "#000000");
    lineColor = color("lineColor", "Couleur du trait pointillé", "#999999");

    name = "milestones";
    displayName = "Jalons";
    slices: FormattingSettingsSlice[] = [this.shape, this.size, this.fill, this.stroke, this.lineColor];
}

/** Palette par défaut pour les sites autres que Lille / Ulis (piste, remplissage). */
export const DEFAULT_PALETTE: Array<[string, string]> = [
    ["#f2a14354", "#d9822bff"],
    ["#a77bca54", "#7d4fa3ff"],
    ["#e0707054", "#c0392bff"],
    ["#4fc1c154", "#1f8f8fff"],
    ["#c9b23c54", "#9c8619ff"],
    ["#8c8c8c54", "#5e5e5eff"]
];

/** Couleurs par site : piste transparente + remplissage opaque. */
export class SiteColorsCard extends FormattingSettingsCard {
    lilleTrack = color("lilleTrack", "Lille - piste", "#58bb5b54");
    lilleFill = color("lilleFill", "Lille - remplissage", "#3aa343ff");
    ulisTrack = color("ulisTrack", "Ulis - piste", "#68a3cb54");
    ulisFill = color("ulisFill", "Ulis - remplissage", "#3f7aabff");
    trackOpacity = num("trackOpacity", "Opacité de la piste si couleur sans alpha (%)", 33, 0, 100);
    palette: Array<[formattingSettings.ColorPicker, formattingSettings.ColorPicker]> = DEFAULT_PALETTE.map((pair, i) => [
        color(`palette${i + 1}Track`, `Autres sites ${i + 1} - piste`, pair[0]),
        color(`palette${i + 1}Fill`, `Autres sites ${i + 1} - remplissage`, pair[1])
    ]);

    name = "siteColors";
    displayName = "Couleurs par site";
    slices: FormattingSettingsSlice[] = [
        this.lilleTrack, this.lilleFill, this.ulisTrack, this.ulisFill,
        ...this.palette.reduce<FormattingSettingsSlice[]>((acc, p) => acc.concat(p), []),
        this.trackOpacity
    ];
}

const scaleItems: powerbi.IEnumMember[] = [
    { value: "0-1", displayName: "0 - 1" },
    { value: "0-100", displayName: "0 - 100" }
];

/** Progression. */
export class ProgressCard extends FormattingSettingsCard {
    scale = new formattingSettings.ItemDropdown({ name: "scale", displayName: "Échelle de progression", items: scaleItems, value: scaleItems[0] });
    showLabel = new formattingSettings.ToggleSwitch({ name: "showLabel", displayName: "Afficher « xx % »", value: true });

    name = "progress";
    displayName = "Progression";
    slices: FormattingSettingsSlice[] = [this.scale, this.showLabel];
}

const positionItems: powerbi.IEnumMember[] = [
    { value: "above", displayName: "Au-dessus" },
    { value: "insideLeft", displayName: "Intérieur gauche" }
];

/** Labels de données (nom du Livrable CPV). */
export class LabelsCard extends FormattingSettingsCard {
    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Afficher les labels", value: true });
    position = new formattingSettings.ItemDropdown({ name: "position", displayName: "Position", items: positionItems, value: positionItems[0] });
    font = new formattingSettings.FontControl({
        name: "font",
        displayName: "Police",
        fontFamily: new formattingSettings.FontPicker({ name: "fontFamily", value: "Segoe UI, wf_segoe-ui_normal, helvetica, arial, sans-serif" }),
        fontSize: new formattingSettings.NumUpDown({ name: "fontSize", value: 10 })
    });
    color = color("color", "Couleur", "#777777");

    name = "labels";
    displayName = "Labels de données";
    topLevelSlice = this.show;
    slices: FormattingSettingsSlice[] = [this.position, this.font, this.color];
}

/** Grille de gauche (colonne Nom). */
export class GridCard extends FormattingSettingsCard {
    font = new formattingSettings.FontControl({
        name: "font",
        displayName: "Police",
        fontFamily: new formattingSettings.FontPicker({ name: "fontFamily", value: "Segoe UI, wf_segoe-ui_normal, helvetica, arial, sans-serif" }),
        fontSize: new formattingSettings.NumUpDown({ name: "fontSize", value: 11 })
    });
    textColor = color("textColor", "Couleur du texte", "#333333");
    siteBackground = color("siteBackground", "Fond des lignes Site", "#f3f3f3");

    name = "grid";
    displayName = "Grille";
    slices: FormattingSettingsSlice[] = [this.font, this.textColor, this.siteBackground];
}

const dateFormatItems: powerbi.IEnumMember[] = [
    { value: "DD-MM-YYYY", displayName: "JJ-MM-AAAA" },
    { value: "DD/MM/YYYY", displayName: "JJ/MM/AAAA" },
    { value: "YYYY-MM-DD", displayName: "AAAA-MM-JJ" }
];

export const HIERARCHY_ITEMS: powerbi.IEnumMember[] = HIERARCHIES.map(h => ({ value: h.value, displayName: h.displayName }));

/** En-têtes de la timeline, hiérarchie de dates et format de date. */
export class TimelineCard extends FormattingSettingsCard {
    hierarchy = new formattingSettings.ItemDropdown({ name: "hierarchy", displayName: "Hiérarchie de dates", items: HIERARCHY_ITEMS, value: HIERARCHY_ITEMS[2] });
    yearBackground = color("yearBackground", "Fond - niveau haut", "#e9eef3");
    yearTextColor = color("yearTextColor", "Texte - niveau haut", "#333333");
    monthBackground = color("monthBackground", "Fond - niveau bas", "#f7f9fb");
    monthTextColor = color("monthTextColor", "Texte - niveau bas", "#555555");
    gridLineColor = color("gridLineColor", "Lignes de grille", "#e6e6e6");
    dateFormat = new formattingSettings.ItemDropdown({ name: "dateFormat", displayName: "Format de date (infobulles)", items: dateFormatItems, value: dateFormatItems[0] });

    name = "timeline";
    displayName = "Timeline";
    slices: FormattingSettingsSlice[] = [this.hierarchy, this.yearBackground, this.yearTextColor, this.monthBackground, this.monthTextColor, this.gridLineColor, this.dateFormat];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    layout = new LayoutCard();
    milestones = new MilestoneCard();
    siteColors = new SiteColorsCard();
    progress = new ProgressCard();
    labels = new LabelsCard();
    grid = new GridCard();
    timeline = new TimelineCard();

    cards = [this.layout, this.milestones, this.siteColors, this.progress, this.labels, this.grid, this.timeline];
}
