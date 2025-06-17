// GridDividerID_v6_Grouped.jsx
// InDesign用 グリッド分割スクリプト (グループ化対応版)

#targetengine "session"; // スクリプトエンジンを指定

// --- グローバル変数・設定 ---
var SCRIPT_NAME = "グリッド分割";
var DECIMAL_PLACES = 3; // 小数点以下の表示桁数 (切り捨て)

// --- メイン処理 ---
function main() {
    // ドキュメントが開かれているか確認
    if (app.documents.length === 0) {
        alert("処理対象のドキュメントが開かれていません。", SCRIPT_NAME, true);
        return;
    }
    var doc = app.activeDocument;

    // オブジェクトが選択されているか確認
    if (doc.selection.length === 0) {
        alert("分割する四角形オブジェクトを選択してください。", SCRIPT_NAME, true);
        return;
    }
    if (doc.selection.length > 1) {
        alert("四角形オブジェクトは1つだけ選択してください。", SCRIPT_NAME, true);
        return;
    }

    var selectedItem = doc.selection[0];
    var itemClassName = ""; // エラー表示用に初期化

    try {
        itemClassName = selectedItem.constructor.name;
    } catch(e){}


    // 選択されたオブジェクトが四角形(Rectangle)であるか確認
    if (!(selectedItem instanceof Rectangle)) {
        alert("四角形オブジェクトを選択してください。\n(現在選択されているオブジェクト: " + itemClassName + ")", SCRIPT_NAME, true);
        return;
    }

    // オブジェクトが空かどうかの検証
    if (!isEmptyRectangle(selectedItem, doc)) {
        alert("選択された四角形オブジェクトは空である必要があります。\n（内部にコンテンツがなく、塗りと線が「なし」で線幅が0の状態）", SCRIPT_NAME, true);
        return;
    }

    // UI表示と値の取得
    var dialogResult = showGridDialog(selectedItem, doc);

    if (dialogResult === null) {
        // キャンセルされた場合
        return;
    }

    // UIから取得した値
    var rowCount = parseInt(dialogResult.rowCount) || 1;
    var rowGutterPt = convertMmToPt(parseFloat(dialogResult.rowGutter) || 0);
    var colCount = parseInt(dialogResult.colCount) || 1;
    var colGutterPt = convertMmToPt(parseFloat(dialogResult.colGutter) || 0);

    // Undoのための処理グループ化
    app.doScript(
        function() {
            processGridDivision(doc, selectedItem, rowCount, rowGutterPt, colCount, colGutterPt);
        },
        ScriptLanguage.JAVASCRIPT,
        [],
        UndoModes.ENTIRE_SCRIPT,
        SCRIPT_NAME + " 処理"
    );
}

// --- 「なし」スウォッチを取得するヘルパー関数 ---
function getNoneSwatch(document) {
    var noneSwatch;
    var noneNameID = "$ID/None"; // InDesign内部名 (推奨)
    var noneNameEN = "None";     // 英語環境
    var noneNameJA = "[なし]";   // 日本語環境の一般的な名前

    if (document.swatches.itemByName(noneNameID).isValid) {
        noneSwatch = document.swatches.itemByName(noneNameID);
    } else if (document.swatches.itemByName(noneNameEN).isValid) {
        noneSwatch = document.swatches.itemByName(noneNameEN);
    } else if (document.swatches.itemByName(noneNameJA).isValid) {
        noneSwatch = document.swatches.itemByName(noneNameJA);
    } else {
        $.writeln("警告: 標準の「なし」スウォッチ('$ID/None', 'None', '[なし]')が見つかりません。fillColor/strokeColorに文字列'None'を使用します。");
        return "None"; // スウォッチオブジェクトではなく文字列"None"を返すことで、適用時にInDesignが解釈することを期待
    }
    return noneSwatch; // Colorオブジェクト(スウォッチ)を返す
}


// --- 空の四角形オブジェクトか検証する関数 ---
function isEmptyRectangle(item, document) {
    if (!(item instanceof Rectangle)) return false;

    // 1. グラフィックコンテンツ（画像など）がないか
    if (item.graphics.length > 0) {
        $.writeln("検証NG: 画像あり");
        return false;
    }

    var noneSwatchObject = getNoneSwatch(document);
    var noneSwatchName = (typeof noneSwatchObject === 'string') ? noneSwatchObject : noneSwatchObject.name;

    // 2. 塗りが「なし」か
    // スウォッチで[なし]が設定されている場合(.name)と、色が何も適用されていない状態(.constructor.nameが'NoColor')の両方を考慮
    if (item.fillColor.name !== noneSwatchName && item.fillColor.constructor.name !== "NoColor" ) {
        if (item.fillColor.name !== "None" && item.fillColor.name !== "[なし]" && item.fillColor.name !== "$ID/None") {
             $.writeln("検証NG: 塗りが設定されています (FillColor Name: " + item.fillColor.name + ", Constructor: " + item.fillColor.constructor.name + ")");
             return false;
        }
    }

    // 3. 線が「なし」か (線幅が0 または 線色が「なし」)
    if (item.strokeWeight > 0) {
        if (item.strokeColor.name !== noneSwatchName && item.strokeColor.constructor.name !== "NoColor") {
            if (item.strokeColor.name !== "None" && item.strokeColor.name !== "[なし]" && item.strokeColor.name !== "$ID/None") {
                $.writeln("検証NG: 線が設定されています (StrokeColor Name: " + item.strokeColor.name + ", Constructor: " + item.strokeColor.constructor.name + ", Weight: " + item.strokeWeight + ")");
                return false;
            }
        }
    }
    return true;
}


// --- UIダイアログ表示とインタラクション ---
function showGridDialog(selectedItem, doc) {
    var dlg = new Window("dialog", SCRIPT_NAME + "設定");
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    dlg.margins = 15;
    dlg.spacing = 10;

    var bounds = selectedItem.geometricBounds;
    var objInitialHeightPt = bounds[2] - bounds[0];
    var objInitialWidthPt = bounds[3] - bounds[1];

    // --- 行設定 ---
    var rowPanel = dlg.add("panel", undefined, "行");
    rowPanel.orientation = "column"; rowPanel.alignChildren = ["fill", "top"]; rowPanel.margins = [10,15,10,10]; rowPanel.spacing = 8;
        var rowCountGroup = rowPanel.add("group"); rowCountGroup.alignment = ["fill", "center"];
        rowCountGroup.add("statictext", undefined, "段数:");
        var rowCountEdit = rowCountGroup.add("edittext", undefined, "2"); rowCountEdit.preferredSize.width = 60;
        var rowHeightGroup = rowPanel.add("group"); rowHeightGroup.alignment = ["fill", "center"];
        rowHeightGroup.add("statictext", undefined, "高さ(1枠):");
        var rowHeightVal = rowHeightGroup.add("statictext", [0,0,80,20], "--- mm", {readonly:true});
        var rowGutterGroup = rowPanel.add("group"); rowGutterGroup.alignment = ["fill", "center"];
        rowGutterGroup.add("statictext", undefined, "間隔:");
        var rowGutterEdit = rowGutterGroup.add("edittext", undefined, "5"); rowGutterEdit.preferredSize.width = 60;
        rowGutterGroup.add("statictext", undefined, "mm");
        var rowTotalGroup = rowPanel.add("group"); rowTotalGroup.alignment = ["fill", "center"];
        rowTotalGroup.add("statictext", undefined, "合計:");
        var rowTotalVal = rowTotalGroup.add("statictext", [0,0,80,20], roundValue(convertPtToMm(objInitialHeightPt), DECIMAL_PLACES) + " mm", {readonly:true});

    // --- 列設定 ---
    var colPanel = dlg.add("panel", undefined, "列");
    colPanel.orientation = "column"; colPanel.alignChildren = ["fill", "top"]; colPanel.margins = [10,15,10,10]; colPanel.spacing = 8;
        var colCountGroup = colPanel.add("group"); colCountGroup.alignment = ["fill", "center"];
        colCountGroup.add("statictext", undefined, "段数:");
        var colCountEdit = colCountGroup.add("edittext", undefined, "2"); colCountEdit.preferredSize.width = 60;
        var colWidthGroup = colPanel.add("group"); colWidthGroup.alignment = ["fill", "center"];
        colWidthGroup.add("statictext", undefined, "幅(1枠):");
        var colWidthVal = colWidthGroup.add("statictext", [0,0,80,20], "--- mm", {readonly:true});
        var colGutterGroup = colPanel.add("group"); colGutterGroup.alignment = ["fill", "center"];
        colGutterGroup.add("statictext", undefined, "間隔:");
        var colGutterEdit = colGutterGroup.add("edittext", undefined, "5");
        colGutterEdit.preferredSize.width = 60;
        colGutterGroup.add("statictext", undefined, "mm");
        var colTotalGroup = colPanel.add("group"); colTotalGroup.alignment = ["fill", "center"];
        colTotalGroup.add("statictext", undefined, "合計:");
        var colTotalVal = colTotalGroup.add("statictext", [0,0,80,20], roundValue(convertPtToMm(objInitialWidthPt), DECIMAL_PLACES) + " mm", {readonly:true});

    // --- ボタン ---
    var buttonGroup = dlg.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = ["right", "center"];
    var okButton = buttonGroup.add("button", undefined, "OK", { name: "ok" });
    var cancelButton = buttonGroup.add("button", undefined, "キャンセル", { name: "cancel" });

    // --- UI動的更新ロジック ---
    function updateFrameSizes() {
        try {
            var rCount = parseInt(rowCountEdit.text) || 1;
            var rGutterMm = parseFloat(rowGutterEdit.text) || 0;
            var rGutterPt = convertMmToPt(rGutterMm);

            var cCount = parseInt(colCountEdit.text) || 1;
            var cGutterMm = parseFloat(colGutterEdit.text) || 0;
            var cGutterPt = convertMmToPt(cGutterMm);

            var frameHeightPt = (objInitialHeightPt - (rCount - 1) * rGutterPt) / rCount;
            var frameWidthPt = (objInitialWidthPt - (cCount - 1) * cGutterPt) / cCount;

            rowHeightVal.text = (frameHeightPt > 0.0001) ? roundValue(convertPtToMm(frameHeightPt), DECIMAL_PLACES) + " mm" : "エラー";
            colWidthVal.text = (frameWidthPt > 0.0001) ? roundValue(convertPtToMm(frameWidthPt), DECIMAL_PLACES) + " mm" : "エラー";
        } catch (e) {
            rowHeightVal.text = "計算エラー";
            colWidthVal.text = "計算エラー";
        }
    }

    updateFrameSizes();
    rowCountEdit.onChanging = updateFrameSizes;
    rowGutterEdit.onChanging = updateFrameSizes;
    colCountEdit.onChanging = updateFrameSizes;
    colGutterEdit.onChanging = updateFrameSizes;


    if (dlg.show() == 1) { // OKボタンが押された
        var rCountFinal = parseInt(rowCountEdit.text);
        var rGutterFinal = parseFloat(rowGutterEdit.text);
        var cCountFinal = parseInt(colCountEdit.text);
        var cGutterFinal = parseFloat(colGutterEdit.text);

        if (isNaN(rCountFinal) || rCountFinal <= 0 || isNaN(rGutterFinal) || rGutterFinal < 0 ||
            isNaN(cCountFinal) || cCountFinal <= 0 || isNaN(cGutterFinal) || cGutterFinal < 0) {
            alert("入力値が無効です。数値で0以上の値を入力してください（段数は1以上）。", SCRIPT_NAME, true);
            return null;
        }
        var finalFrameHeightPt = (objInitialHeightPt - (rCountFinal - 1) * convertMmToPt(rGutterFinal)) / rCountFinal;
        var finalFrameWidthPt = (objInitialWidthPt - (cCountFinal - 1) * convertMmToPt(cGutterFinal)) / cCountFinal;

        if (finalFrameHeightPt <= 0.0001 || finalFrameWidthPt <= 0.0001) {
            alert("計算された1枠あたりのサイズが小さすぎます（ほぼ0またはマイナス）。\n段数または間隔を調整してください。", SCRIPT_NAME, true);
            return null;
        }
        return {
            rowCount: rowCountEdit.text,
            rowGutter: rowGutterEdit.text,
            colCount: colCountEdit.text,
            colGutter: colGutterEdit.text
        };
    } else { // キャンセルまたは閉じられた
        return null;
    }
}


// --- グリッド分割処理 ---
function processGridDivision(doc, item, rowCount, rowGutterPt, colCount, colGutterPt) {
    // ★変更点: 作成したフレームを格納するための配列を準備
    var createdFrames = [];
    
    var itemPage = item.parentPage || doc.activeLayer.parent;
    var originalHUnits = doc.viewPreferences.horizontalMeasurementUnits;
    var originalVUnits = doc.viewPreferences.verticalMeasurementUnits;
    var hUnitsChanged = false, vUnitsChanged = false;
    var noneSwatchString = "None";

    try {
        if (originalHUnits !== MeasurementUnits.POINTS) { doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.POINTS; hUnitsChanged = true; }
        if (originalVUnits !== MeasurementUnits.POINTS) { doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS; vUnitsChanged = true; }

        var bounds = item.geometricBounds;
        var itemY1 = bounds[0]; var itemX1 = bounds[1];
        var itemHeight = bounds[2] - bounds[0]; var itemWidth = bounds[3] - bounds[1];

        var frameHeight = (itemHeight - (rowCount - 1) * rowGutterPt) / rowCount;
        var frameWidth = (itemWidth - (colCount - 1) * colGutterPt) / colCount;

        if (frameHeight <= 0.0001 || frameWidth <= 0.0001) {
            throw new Error("E001: 分割後の1枠あたりのサイズが小さすぎます。");
        }

        var originalContentType = item.contentType;
        var originalItemLayer = item.itemLayer;

        var currentY = itemY1;
        for (var r = 0; r < rowCount; r++) {
            var currentX = itemX1;
            for (var c = 0; c < colCount; c++) {
                var newFrameBounds = [currentY, currentX, currentY + frameHeight, currentX + frameWidth];
                
                var newFrame;
                try {
                    newFrame = itemPage.rectangles.add({
                        itemLayer: originalItemLayer,
                        geometricBounds: newFrameBounds,
                        strokeWeight: 0,
                        fillColor: noneSwatchString,
                        strokeColor: noneSwatchString
                    });
                     // ★変更点: 作成したフレームを配列に追加
                    createdFrames.push(newFrame);
                } catch (e_addRect) {
                    throw new Error("E002: 新規フレームの作成に失敗しました。Bounds: " + newFrameBounds.join(", ") + "\nOriginalError: " + e_addRect.message);
                }

                if (!newFrame.isValid) {
                    throw new Error("E003: 作成されたフレームが無効です (行:"+r+", 列:"+c+")。");
                }

                if (originalContentType === ContentType.GRAPHIC_TYPE) {
                    try {
                        newFrame.contentType = ContentType.GRAPHIC_TYPE;
                    } catch (e_contentType) {
                         $.writeln("contentType設定でエラー: " + e_contentType.message);
                    }
                }
                currentX += frameWidth + colGutterPt;
            }
            currentY += frameHeight + rowGutterPt;
        }
        item.remove();
        
        // ★変更点: 作成されたフレームが複数あればグループ化する
        if (createdFrames.length > 1) {
            try {
                itemPage.groups.add(createdFrames);
            } catch(e_group) {
                // グループ化に失敗しても処理は続行するが、コンソールにログは残す
                $.writeln("作成されたフレームのグループ化に失敗しました: " + e_group.message);
            }
        }

    } catch (e_process_internal) {
        var errorMessage = "分割処理中に内部エラーが発生しました。\n";
        errorMessage += "メッセージ: " + e_process_internal.message;
        if (e_process_internal.line) errorMessage += "\nライン: " + e_process_internal.line;
        $.writeln(errorMessage.replace(/\n/g, " || "));
        throw new Error(errorMessage);
    } finally {
        if (hUnitsChanged) doc.viewPreferences.horizontalMeasurementUnits = originalHUnits;
        if (vUnitsChanged) doc.viewPreferences.verticalMeasurementUnits = originalVUnits;
    }
}

// --- ヘルパー関数 ---
function convertMmToPt(mmValue) {
    return (parseFloat(mmValue) / 25.4) * 72;
}
function convertPtToMm(ptValue) {
    return (parseFloat(ptValue) / 72) * 25.4;
}
function roundValue(value, digits) {
    var multiplier = Math.pow(10, digits || 0);
    return Math.floor(parseFloat(value) * multiplier) / multiplier;
}

// --- スクリプト実行 ---
try {
    main();
} catch (e_global) {
    var finalMessage = SCRIPT_NAME + "実行中に予期せぬエラーが発生しました:\n";
    if (e_global instanceof Error && typeof e_global.message !== 'undefined') {
        finalMessage += e_global.message;
    } else {
        finalMessage += String(e_global);
    }
    alert(finalMessage, SCRIPT_NAME, true);
}
