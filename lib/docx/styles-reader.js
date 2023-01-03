exports.readStylesXml = readStylesXml;
exports.Styles = Styles;
exports.defaultStyles = new Styles({}, {});

function Styles(paragraphStyles, characterStyles, tableStyles, numberingStyles, defaultStyles) {
    return {
        findParagraphStyleById: function(styleId) {
            return paragraphStyles[styleId];
        },
        findCharacterStyleById: function(styleId) {
            return characterStyles[styleId];
        },
        findTableStyleById: function(styleId) {
            return tableStyles[styleId];
        },
        findNumberingStyleById: function(styleId) {
            return numberingStyles[styleId];
        },
        getDefault: function() {
            return defaultStyles;
        }
    };
}

Styles.EMPTY = new Styles({}, {}, {}, {});

function readStylesXml(root) {
    var paragraphStyles = {};
    var characterStyles = {};
    var tableStyles = {};
    var numberingStyles = {};
    var defaultStyles = {};

    var styles = {
        "paragraph": paragraphStyles,
        "character": characterStyles,
        "table": tableStyles,
        default: defaultStyles
    };

    root.getElementsByTagName("w:docDefaults").forEach(function(styleElement) {
        styleElement.getElementsByTagName("w:rPrDefault").forEach(function(rPrDefault) {
            var element = rPrDefault.firstOrEmpty("w:rPr");
            var fontSizeString = element.firstOrEmpty("w:sz").attributes["w:val"];
            // w:sz gives the font size in half points, so halve the value to get the size in points
            var fontSize = /^[0-9]+$/.test(fontSizeString)
                ? parseInt(fontSizeString, 10) / 2
                : null;
            styles.default.runStyle = {
                font: element.firstOrEmpty("w:rFonts").attributes["w:ascii"],
                ...(fontSize ? {"fontSize": fontSize.toString()} : {}),
            }
        })});

    root.getElementsByTagName("w:style").forEach(function(styleElement) {
        var style = readStyleElement(styleElement);

        if (style.type === "numbering") {
            numberingStyles[style.styleId] = readNumberingStyleElement(styleElement);
        } else {
            var styleSet = styles[style.type];
            if (styleSet) {
                styleSet[style.styleId] = style;
                styleSet[style.styleId].runStyle = {...styles.default.runStyle, ...style.runStyle};
            }
        }
    });
    // recursively resolve basedOn styles
    for (var styleId in paragraphStyles) {
        paragraphStyles[styleId].runStyle = resolveBasedOnStyle(paragraphStyles, styleId);
    }
    return new Styles(paragraphStyles, characterStyles, tableStyles, numberingStyles, defaultStyles);
}

function resolveBasedOnStyle(styles, styleId) {
    var style = styles[styleId];
    let currentRunStyle = style.runStyle;
    if (style.basedOn) {
        var basedOnStyle = styles[style.basedOn];
        if (basedOnStyle) {
            currentRunStyle = {...resolveBasedOnStyle(styles, style.basedOn), ...basedOnStyle.runStyle, ...currentRunStyle};
        }
    }
    return currentRunStyle;
}

function readBooleanElement(element) {
    if (element) {
        var value = element.attributes["w:val"];
        return value !== "false" && value !== "0";
    } else {
        return false;
    }
}

function readUnderline(element) {
    if (element) {
        var value = element.attributes["w:val"];
        return value !== "false" && value !== "0" && value !== "none";
    } else {
        return false;
    }
}


function readStyleElement(styleElement) {
    var type = styleElement.attributes["w:type"];
    var styleId = styleElement.attributes["w:styleId"];
    var element = styleElement.firstOrEmpty("w:rPr");
    var basedOn = styleElement.firstOrEmpty("w:basedOn").attributes["w:val"];

    var fontSizeString = element.firstOrEmpty("w:sz").attributes["w:val"];
    // w:sz gives the font size in half points, so halve the value to get the size in points
    var fontSize = /^[0-9]+$/.test(fontSizeString)
        ? parseInt(fontSizeString, 10) / 2
        : null;

    
    var runStyle = {
        font: element.firstOrEmpty("w:rFonts").attributes["w:ascii"],
        ...(fontSize ? {"fontSize": fontSize.toString()} : {}),
        color: element.firstOrEmpty("w:color").attributes["w:val"],
        isBold: readBooleanElement(element.first("w:b")),
        alignment: element.firstOrEmpty("w:jc").attributes["w:val"],
        isUnderline: readUnderline(element.first("w:u")),
        spacing: element.firstOrEmpty("w:spacing").attributes["w:line"],
        beforeSpacing: element.firstOrEmpty("w:spacing").attributes["w:before"],
        afterSpacing: element.firstOrEmpty("w:spacing").attributes["w:after"],
        isItalic: readBooleanElement(element.first("w:i")),
        isStrikethrough: readBooleanElement(element.first("w:strike")),
        isAllCaps: readBooleanElement(element.first("w:caps")),
        isSmallCaps: readBooleanElement(element.first("w:smallCaps")),
    };
    for (var key in runStyle) {
        if (!runStyle[key]) {
            delete runStyle[key];
        }
    }
    var name = styleName(styleElement);
    return {type: type, styleId: styleId, basedOn, name: name, runStyle};
}

function styleName(styleElement) {
    var nameElement = styleElement.first("w:name");
    return nameElement ? nameElement.attributes["w:val"] : null;
}

function readNumberingStyleElement(styleElement) {
    var numId = styleElement
        .firstOrEmpty("w:pPr")
        .firstOrEmpty("w:numPr")
        .firstOrEmpty("w:numId")
        .attributes["w:val"];
    return {numId: numId};
}
