//! Drawing objects keep their names, ids, hidden flags, macros and stacking order on save, so a
//! macro-enabled workbook's buttons stay bound and its hidden objects stay hidden.

use gridcraft_model::{Shape, ShapeKind};

use super::minimal;
use crate::package::Package;
use crate::xml::El;
use crate::{read_xlsx, write_xlsx};

const CHART: &str = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:val><c:numRef><c:f>Data!$A$1:$A$3</c:f></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;

fn anchor(col: u32, body: &str) -> String {
    format!(
        "<xdr:twoCellAnchor><xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>{}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>{body}<xdr:clientData/></xdr:twoCellAnchor>",
        col + 2
    )
}

fn shape(id: u32, name: &str, extra: &str, mac: &str) -> String {
    format!(
        r#"<xdr:sp macro="{mac}" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{id}" name="{name}"{extra}/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>"#
    )
}

/// A one-sheet workbook whose drawing holds, bottom to top: a label, a button bound to a macro,
/// a hidden chart and a hidden shape over it.
fn book() -> Vec<u8> {
    let objects = [
        anchor(0, &shape(9, "statusLabel", "", "")),
        anchor(1, &shape(10, "btnRun", "", "Module1.RunUpdate")),
        anchor(
            2,
            r#"<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="40" name="TrendChart" hidden="1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame>"#,
        ),
        anchor(3, &shape(11, "overlay", r#" hidden="1""#, "")),
    ]
    .concat();
    let drawing = format!(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{objects}</xdr:wsDr>"#
    );
    let drels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#;
    let srels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#;
    minimal(
        r#"<sheetData/><drawing r:id="rId1"/>"#,
        &[
            ("xl/worksheets/_rels/sheet1.xml.rels", srels),
            ("xl/drawings/drawing1.xml", &drawing),
            ("xl/drawings/_rels/drawing1.xml.rels", drels),
            ("xl/charts/chart1.xml", CHART),
        ],
        "",
        "",
    )
}

/// (id, name, hidden, macro) of each object of a written drawing, in file order.
fn objects(bytes: &[u8]) -> Vec<(String, String, bool, String)> {
    let mut pkg = Package::open(bytes).unwrap();
    let x = pkg.read_xml("xl/drawings/drawing1.xml").unwrap().unwrap();
    let objs: Vec<&El> =
        x.children.iter().filter_map(|a| a.children.iter().find(|o| matches!(o.name.as_str(), "sp" | "graphicFrame" | "pic"))).collect();
    objs.iter()
        .map(|o| {
            let nv = ["nvSpPr", "nvGraphicFramePr", "nvPicPr"].iter().find_map(|n| o.path(&[n, "cNvPr"])).unwrap();
            (
                nv.attr("id").unwrap().to_string(),
                nv.attr("name").unwrap().to_string(),
                nv.flag("hidden", false),
                o.attr("macro").unwrap_or("").to_string(),
            )
        })
        .collect()
}

fn row(id: &str, name: &str, hidden: bool, mac: &str) -> (String, String, bool, String) {
    (id.into(), name.into(), hidden, mac.into())
}

#[test]
fn names_ids_hidden_flags_macros_and_order_survive_save() {
    let (wb, _) = read_xlsx(&book()).unwrap();
    let s = wb.sheet(0).unwrap();
    let names: Vec<Option<&str>> = s.shapes.iter().map(|o| o.info.name.as_deref()).collect();
    assert_eq!(names, [Some("statusLabel"), Some("btnRun"), Some("overlay")]);
    assert_eq!(s.shapes[1].info.macro_name.as_deref(), Some("Module1.RunUpdate"));
    assert!(s.shapes[2].info.hidden && !s.shapes[0].info.hidden);
    assert_eq!(s.charts[0].info.name.as_deref(), Some("TrendChart"));
    assert_eq!(s.charts[0].info.file_id, Some(40));
    assert!(s.charts[0].info.hidden);

    let saved = write_xlsx(&wb).unwrap();
    assert_eq!(
        objects(&saved),
        [
            row("9", "statusLabel", false, ""),
            row("10", "btnRun", false, "Module1.RunUpdate"),
            row("40", "TrendChart", true, ""),
            row("11", "overlay", true, ""),
        ],
        "the shapes are not all written above the chart"
    );
    let (back, _) = read_xlsx(&saved).unwrap();
    assert_eq!(back.sheet(0).unwrap().shapes, s.shapes);
    assert_eq!(back.sheet(0).unwrap().charts[0].info, s.charts[0].info);
}

#[test]
fn new_and_copied_objects_get_unused_ids_and_go_on_top() {
    let (mut wb, _) = read_xlsx(&book()).unwrap();
    let s = wb.sheet_mut(0).unwrap();
    // A copy of the button keeps its name and macro (as Excel's copy does) but not its id or its
    // place: it goes on top.
    let copy = Shape { id: 900, ..s.shapes[1].clone() };
    let new = Shape { id: 901, kind: ShapeKind::Ellipse, ..Shape { info: Default::default(), ..s.shapes[0].clone() } };
    s.shapes.push(copy);
    s.shapes.push(new);
    let objs = objects(&write_xlsx(&wb).unwrap());
    let ids: Vec<&str> = objs.iter().map(|o| o.0.as_str()).collect();
    assert_eq!(ids, ["9", "10", "40", "11", "2", "3"], "{objs:?}");
    assert_eq!(objs[4], row("2", "btnRun", false, "Module1.RunUpdate"));
    assert_eq!(objs[5].1, "Shape 6");
}

#[test]
fn hostile_names_and_macros_are_escaped_and_capped() {
    let long = "x".repeat(5000);
    let (mut wb, _) = read_xlsx(&book()).unwrap();
    let s = wb.sheet_mut(0).unwrap();
    s.shapes[0].info.name = Some("a\"<b>&".into());
    s.shapes[1].info.macro_name = Some(long.clone());
    let saved = write_xlsx(&wb).unwrap();
    let objs = objects(&saved);
    assert_eq!(objs[0].1, "a\"<b>&");
    let (back, _) = read_xlsx(&saved).unwrap();
    assert_eq!(back.sheet(0).unwrap().shapes[1].info.macro_name.as_deref().map(str::len), Some(1024));
}
