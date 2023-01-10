var _ = require("underscore");

exports.readNumberingXml = readNumberingXml;
exports.Numbering = Numbering;
exports.defaultNumbering = new Numbering({}, {});

function Numbering(nums, abstractNums, styles) {
    var allLevels = _.flatten(
        _.values(abstractNums).map(function (abstractNum) {
            return _.values(abstractNum.levels);
        })
    );

    var levelsByParagraphStyleId = _.indexBy(
        allLevels.filter(function (level) {
            return level.paragraphStyleId != null;
        }),
        "paragraphStyleId"
    );


    function findLevel(numId, level) {
        var num = nums[numId];
        if (num) {
            const numLevel = num[level];
            var abstractNum = abstractNums[num.abstractNumId];
            if (!abstractNum) {
                return null;
            } 
            let subNum;
            if (abstractNum.numStyleLink) {
                const style = styles.findNumberingStyleById(
                    abstractNum.numStyleLink
                );
                subNum = findLevel(style.numId, level);
                subNum = subNum && Object.fromEntries(Object.entries(subNum).filter(([_, v]) => v != null));
            }
            const finalNum = {
                ...subNum,
                ...(abstractNum.levels[level] && Object.fromEntries(Object.entries(abstractNum.levels[level]).filter(([_, v]) => v != null))),
                ...(numLevel && Object.fromEntries(Object.entries(numLevel).filter(([_, v]) => v != null))),
                ...(num.startOverride && {startOverride: num.startOverride}),
            }
            return finalNum;
        } else {
            return null;
        }
    }

    function findLevelByParagraphStyleId(styleId) {
        return levelsByParagraphStyleId[styleId] || null;
    }

    return {
        findLevel: findLevel,
        findLevelByParagraphStyleId: findLevelByParagraphStyleId,
    };
}

function readNumberingXml(root, options) {
    if (!options || !options.styles) {
        throw new Error("styles is missing");
    }

    var abstractNums = readAbstractNums(root);
    var nums = readNums(root, abstractNums);
    return new Numbering(nums, abstractNums, options.styles);
}

function readAbstractNums(root) {
    var abstractNums = {};
    root.getElementsByTagName("w:abstractNum").forEach(function (element) {
        var id = element.attributes["w:abstractNumId"];
        abstractNums[id] = readAbstractNum(element);
    });
    return abstractNums;
}

function readAbstractNum(element) {
    var levels = {};
    var abstractNumId = element.attributes["w:abstractNumId"];
    element.getElementsByTagName("w:lvl").forEach(function (levelElement) {
        var levelIndex = levelElement.attributes["w:ilvl"];
        var numFmt = levelElement.first("w:numFmt").attributes["w:val"];
        var start = levelElement.first("w:start")
            ? levelElement.first("w:start").attributes["w:val"]
            : undefined;
        if (!start) {
            start = 1;
        }
        var lvlText = levelElement.first("w:lvlText").attributes["w:val"];
        var paragraphStyleId =
            levelElement.firstOrEmpty("w:pStyle").attributes["w:val"];

        levels[levelIndex] = {
            start: start,
            abstractNumId: abstractNumId,
            numFmt: numFmt,
            lvlText: lvlText,
            isOrdered: numFmt !== "bullet",
            level: levelIndex,
            paragraphStyleId: paragraphStyleId,
        };
    });

    var numStyleLink =
        element.firstOrEmpty("w:numStyleLink").attributes["w:val"];

    return { levels: levels, numStyleLink: numStyleLink };
}

function readNums(root) {
    var nums = {};
    root.getElementsByTagName("w:num").forEach(function (element) {
        var numId = element.attributes["w:numId"];
        var abstractNumId =
            element.first("w:abstractNumId").attributes["w:val"];
        nums[numId] = { abstractNumId: abstractNumId };

        element
            .getElementsByTagName("w:lvlOverride")
            .forEach(function (lvlOverrideElement) {
                const startOverride =
                    lvlOverrideElement.first("w:startOverride")?.attributes[
                        "w:val"
                    ];
                nums[numId].startOverride = startOverride;
                lvlOverrideElement
                    .getElementsByTagName("w:lvl")
                    .forEach(function (lvlElement) {
                        const level = lvlElement.attributes["w:ilvl"];
                        const start = lvlElement.attributes["w:val"];
                        const numFmt =
                            lvlElement.first("w:numFmt")?.attributes["w:val"];
                        const lvlText =
                            lvlElement.first("w:lvlText")?.attributes["w:val"];
                        
                        nums[numId][level] = {
                            start: start,
                            numFmt: numFmt,
                            lvlText: lvlText,
                        };
                    });
            });
    });
    return nums;
}
