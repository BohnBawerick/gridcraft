//! Drawing objects keep their names, ids, hidden flags, macros and stacking order on save, so a
//! macro-enabled workbook's buttons stay bound and its hidden objects stay hidden.

use std::sync::Arc;

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
fn long_names_and_macros_are_escaped_and_preserved() {
    let long = "x\"<b>&".repeat(1000);
    let (mut wb, _) = read_xlsx(&book()).unwrap();
    let s = wb.sheet_mut(0).unwrap();
    s.shapes[0].info.name = Some(long.clone());
    s.shapes[1].info.macro_name = Some(long.clone());
    let saved = write_xlsx(&wb).unwrap();
    let objs = objects(&saved);
    assert_eq!(objs[0].1, long);
    let (back, _) = read_xlsx(&saved).unwrap();
    assert_eq!(back.sheet(0).unwrap().shapes[0].info.name.as_deref(), Some(long.as_str()));
    assert_eq!(back.sheet(0).unwrap().shapes[1].info.macro_name.as_deref(), Some(long.as_str()));
}

const COMPAT: &str = r#"<a:ext uri="63B3BB69-23CF-44e3-9099-C40C66FF867C"><a14:compatExt xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" spid="_x0000_s1025"/></a:ext>"#;

fn legacy_control_object(extension: &str, connector: bool) -> String {
    let (tag, nvtag) = if connector { ("cxnSp", "nvCxnSpPr") } else { ("sp", "nvSpPr") };
    format!(
        r#"<xdr:{tag}><xdr:{nvtag}><xdr:cNvPr id="1025" name="Legacy button" hidden="1"><a:extLst>{extension}</a:extLst></xdr:cNvPr></xdr:{nvtag}><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr></xdr:{tag}>"#
    )
}

fn legacy_control(extension: &str, connector: bool) -> Vec<u8> {
    legacy_control_drawing(&anchor(0, &legacy_control_object(extension, connector)))
}

fn legacy_control_drawing(contents: &str) -> Vec<u8> {
    let drawing = format!(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">{}</xdr:wsDr>"#,
        contents
    );
    let types = super::CT.replace("</Types>", r#"<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/></Types>"#);
    let workbook = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
    let wb_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#;
    let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rId1"/><legacyDrawing r:id="rId2"/><controls><control shapeId="1025" name="Legacy button" r:id="rId3"/></controls></worksheet>"#;
    let sheet_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/></Relationships>"#;
    let vml = r#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="_x0000_s1025"><x:ClientData ObjectType="Button"><x:FmlaMacro>Module1.Run</x:FmlaMacro></x:ClientData></v:shape></xml>"#;
    let control = r#"<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="Button"/>"#;
    super::make_zip(&[
        ("[Content_Types].xml", types.as_bytes()),
        ("_rels/.rels", super::ROOT_RELS.as_bytes()),
        ("xl/workbook.xml", workbook.as_bytes()),
        ("xl/_rels/workbook.xml.rels", wb_rels.as_bytes()),
        ("xl/worksheets/sheet1.xml", sheet.as_bytes()),
        ("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels.as_bytes()),
        ("xl/drawings/drawing1.xml", drawing.as_bytes()),
        ("xl/drawings/_rels/drawing1.xml.rels", br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#),
        ("xl/media/image1.png", &[1, 2, 3]),
        ("xl/drawings/vmlDrawing1.vml", vml.as_bytes()),
        ("xl/ctrlProps/ctrlProp1.xml", control.as_bytes()),
    ])
}

#[test]
fn legacy_control_association_and_relationships_survive_two_saves() {
    for connector in [false, true] {
        for uri in ["63B3BB69-23CF-44e3-9099-C40C66FF867C", "{63B3BB69-23CF-44e3-9099-C40C66FF867C}"] {
            let extension = COMPAT.replace("63B3BB69-23CF-44e3-9099-C40C66FF867C", uri);
            let mut bytes = legacy_control(&extension, connector);
            for _ in 0..2 {
                let (wb, report) = read_xlsx(&bytes).unwrap();
                assert!(!report.warnings.iter().any(|w| w.contains("cNvPr")), "{:?}", report.warnings);
                assert_eq!(wb.sheet(0).unwrap().shapes[0].info.compat_spid.as_deref(), Some("_x0000_s1025"));
                bytes = write_xlsx(&wb).unwrap();
                let mut pkg = Package::open(&bytes).unwrap();
                let drawing = pkg.read_xml("xl/drawings/drawing1.xml").unwrap().unwrap();
                let nv = drawing.children[0].path(&["sp", "nvSpPr", "cNvPr"]).unwrap();
                let ext = nv.path(&["extLst", "ext"]).unwrap();
                assert_eq!(ext.attr("uri"), Some("{63B3BB69-23CF-44e3-9099-C40C66FF867C}"));
                let spid = ext.child("compatExt").unwrap().attr("spid").unwrap();
                let sheet = pkg.read_xml("xl/worksheets/sheet1.xml").unwrap().unwrap();
                let rels = pkg.rels("xl/worksheets/sheet1.xml").unwrap();
                let vml_id = sheet.child("legacyDrawing").unwrap().attr("id").unwrap();
                let vml_part = &rels.iter().find(|r| r.id == vml_id).unwrap().target;
                let vml = pkg.read_xml(vml_part).unwrap().unwrap();
                let vml_shape = vml.child("shape").unwrap();
                assert_eq!(vml_shape.attr("id"), Some(spid));
                assert_eq!(vml_shape.path(&["ClientData", "FmlaMacro"]).unwrap().text, "Module1.Run");
                let control = sheet.path(&["controls", "control"]).unwrap();
                assert_eq!(format!("_x0000_s{}", control.attr("shapeId").unwrap()), spid);
                let control_part = &rels.iter().find(|r| Some(r.id.as_str()) == control.attr("id")).unwrap().target;
                assert_eq!(pkg.read_xml(control_part).unwrap().unwrap().attr("objectType"), Some("Button"));
            }
        }
    }
}

#[test]
fn unsupported_object_extensions_warn_with_the_object_name() {
    for extra in [
        r#"<a:ext uri="urn:synthetic:unknown"><a:unknown/></a:ext>"#,
        r#"<a:ext uri="63B3BB69-23CF-44e3-9099-C40C66FF867C"><a:compatExt/></a:ext>"#,
        COMPAT,
    ] {
        let (wb, report) = read_xlsx(&legacy_control(&format!("{COMPAT}{extra}"), false)).unwrap();
        assert_eq!(wb.sheet(0).unwrap().shapes[0].info.compat_spid.as_deref(), Some("_x0000_s1025"));
        assert!(report.warnings.iter().any(|w| w.contains("Legacy button") && w.contains("unsupported cNvPr extension")), "{:?}", report.warnings);
    }
}

fn properties<'a>(el: &'a El, out: &mut Vec<&'a El>) {
    if el.name == "cNvPr" {
        out.push(el);
    }
    for child in &el.children {
        properties(child, out);
    }
}

#[test]
fn every_serializer_keeps_compatibility_links_and_chartex_visibility() {
    use gridcraft_model::{ChartKind, Color, Image};

    for hidden in [false, true] {
        let (mut wb, _) = read_xlsx(&book()).unwrap();
        let sheet = wb.sheet_mut(0).unwrap();
        let mut shape = sheet.shapes[0].clone();
        shape.info.compat_spid = Some("_x0000_s1025".into());
        shape.info.hidden = hidden;
        sheet.shapes = vec![shape.clone()];
        sheet.images.push(Image {
            id: 99,
            info: shape.info.clone(),
            anchor: shape.anchor,
            data: vec![1, 2, 3],
            mime: "image/png".into(),
            alt: "Picture & description".into(),
        });
        shape.kind = ShapeKind::Ink;
        shape.text = "0,0;10,10".into();
        shape.fill = Color::Rgb(3 << 16);
        shape.line = Color::Rgb(0x123456);
        sheet.shapes.push(shape.clone());
        sheet.charts[0].info = shape.info.clone();
        let mut chart = sheet.charts[0].clone();
        chart.kind = ChartKind::Waterfall;
        sheet.charts.push(chart);
        let (legacy, _) = read_xlsx(&legacy_control(COMPAT, false)).unwrap();
        sheet.preserved = legacy.sheet(0).unwrap().preserved.clone();
        let mut vml = String::from(r#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">"#);
        for (index, info) in sheet
            .images
            .iter_mut()
            .map(|o| &mut o.info)
            .chain(sheet.charts.iter_mut().map(|o| &mut o.info))
            .chain(sheet.shapes.iter_mut().map(|o| &mut o.info))
            .enumerate()
        {
            let id = 1025 + index as u32;
            info.file_id = Some(id);
            info.order = Some(index as u32);
            info.compat_spid = Some(format!("_x0000_s{id}"));
            vml.push_str(&format!(r#"<v:shape id="_x0000_s{id}"><x:ClientData ObjectType="Button"/></v:shape>"#));
        }
        vml.push_str("</xml>");
        Arc::make_mut(sheet.preserved.as_mut().unwrap()).vml.as_mut().unwrap().data = Arc::from(vml.into_bytes());
        let bytes = write_xlsx(&wb).unwrap();
        let mut pkg = Package::open(&bytes).unwrap();
        let drawing = pkg.read_xml("xl/drawings/drawing1.xml").unwrap().unwrap();
        let mut props = vec![];
        properties(&drawing, &mut props);
        assert_eq!(props.len(), 6);
        for prop in props {
            assert_eq!(prop.flag("hidden", false), hidden);
            let spid = format!("_x0000_s{}", prop.attr("id").unwrap());
            assert_eq!(prop.path(&["extLst", "ext", "compatExt"]).unwrap().attr("spid"), Some(spid.as_str()));
        }
        let (back, report) = read_xlsx(&bytes).unwrap();
        assert!(!report.warnings.iter().any(|w| w.contains("cNvPr")), "{:?}", report.warnings);
        let sheet = back.sheet(0).unwrap();
        assert_eq!(sheet.images[0].alt, "Picture & description");
        assert_eq!(sheet.shapes[1].kind, ShapeKind::Ink);
        assert_eq!(sheet.charts[1].kind, ChartKind::Waterfall);
        for info in sheet.shapes.iter().map(|o| &o.info).chain(sheet.images.iter().map(|o| &o.info)).chain(sheet.charts.iter().map(|o| &o.info)) {
            assert_eq!(info.compat_spid, Some(format!("_x0000_s{}", info.file_id.unwrap())));
            assert_eq!(info.hidden, hidden);
        }
        wb.sheet_mut(0).unwrap().preserved = None;
        let saved = write_xlsx(&wb).unwrap();
        let mut pkg = Package::open(&saved).unwrap();
        let drawing = pkg.read_xml("xl/drawings/drawing1.xml").unwrap().unwrap();
        let mut props = vec![];
        properties(&drawing, &mut props);
        assert_eq!(props.len(), 6);
        assert!(props.iter().all(|p| p.path(&["extLst", "ext", "compatExt"]).is_none()));
    }
}

#[test]
fn alternate_content_control_twins_keep_vml_without_creating_objects() {
    let shape = legacy_control_object(COMPAT, false);
    let picture = format!(
        r#"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1025" name="Legacy button" hidden="1"><a:extLst>{COMPAT}</a:extLst></xdr:cNvPr><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdImage"/></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr></xdr:pic>"#
    );
    for object in [shape, picture] {
        for drawing_level in [false, true] {
            let content = if drawing_level { anchor(0, &object) } else { object.clone() };
            let alternate = format!(
                r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">{content}</mc:Choice><mc:Fallback/></mc:AlternateContent>"#
            );
            let drawing = if drawing_level { alternate } else { anchor(0, &alternate) };
            let mut bytes = legacy_control_drawing(&drawing);
            let mut pkg = Package::open(&bytes).unwrap();
            let expected_vml = pkg.read("xl/drawings/vmlDrawing1.vml").unwrap().unwrap();
            let expected_control = pkg.read("xl/ctrlProps/ctrlProp1.xml").unwrap().unwrap();
            for _ in 0..2 {
                let (wb, _) = read_xlsx(&bytes).unwrap();
                let sheet = wb.sheet(0).unwrap();
                assert!(sheet.shapes.is_empty() && sheet.images.is_empty() && sheet.charts.is_empty());
                bytes = write_xlsx(&wb).unwrap();
                let mut pkg = Package::open(&bytes).unwrap();
                let sheet = pkg.read_xml("xl/worksheets/sheet1.xml").unwrap().unwrap();
                assert!(sheet.child("drawing").is_none());
                let rels = pkg.rels("xl/worksheets/sheet1.xml").unwrap();
                let vml_id = sheet.child("legacyDrawing").unwrap().attr("id").unwrap();
                let vml_part = &rels.iter().find(|r| r.id == vml_id).unwrap().target;
                assert_eq!(pkg.read(vml_part).unwrap().unwrap(), expected_vml);
                let control = sheet.path(&["controls", "control"]).unwrap();
                assert_eq!(control.attr("shapeId"), Some("1025"));
                let control_part = &rels.iter().find(|r| Some(r.id.as_str()) == control.attr("id")).unwrap().target;
                assert_eq!(pkg.read(control_part).unwrap().unwrap(), expected_control);
            }
        }
    }
}

fn replace_part(bytes: &[u8], name: &str, replacement: Option<&[u8]>) -> Vec<u8> {
    use std::io::Read;

    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut parts = vec![];
    for index in 0..zip.len() {
        let mut part = zip.by_index(index).unwrap();
        let mut data = vec![];
        if part.name() == name {
            let Some(replacement) = replacement else { continue };
            data.extend_from_slice(replacement);
        } else {
            part.read_to_end(&mut data).unwrap();
        }
        parts.push((part.name().to_owned(), data));
    }
    super::make_zip(&parts.iter().map(|(n, d)| (n.as_str(), d.as_slice())).collect::<Vec<_>>())
}

#[test]
fn discarded_vml_clears_compatibility_links_and_names_the_object() {
    for part in ["xl/drawings/vmlDrawing1.vml", "xl/ctrlProps/ctrlProp1.xml"] {
        let bytes = replace_part(&legacy_control(COMPAT, false), part, None);
        let (wb, report) = read_xlsx(&bytes).unwrap();
        let sheet = wb.sheet(0).unwrap();
        assert!(sheet.preserved.as_ref().and_then(|p| p.vml.as_ref()).is_none());
        assert!(sheet.shapes[0].info.compat_spid.is_none());
        assert!(report.warnings.iter().any(|w| w.contains("Legacy button") && w.contains("compatibility link removed")), "{:?}", report.warnings);
        let bytes = write_xlsx(&wb).unwrap();
        let mut pkg = Package::open(&bytes).unwrap();
        let drawing = pkg.read_xml("xl/drawings/drawing1.xml").unwrap().unwrap();
        assert!(drawing.children[0].path(&["sp", "nvSpPr", "cNvPr", "extLst"]).is_none());
    }
}

fn note() -> gridcraft_model::Comment {
    gridcraft_model::Comment {
        author: "Author".into(),
        text: "A synthetic note".into(),
        replies: vec![],
        threaded: false,
        resolved: false,
        visible: true,
        dates: vec![],
    }
}

#[test]
fn note_ids_cannot_replace_missing_legacy_shapes() {
    for backing in ["kept", "missing", "collision", "unmergeable"] {
        let (mut wb, _) = read_xlsx(&legacy_control(COMPAT, false)).unwrap();
        let sheet = wb.sheet_mut(0).unwrap();
        sheet.comments.insert(gridcraft_core::CellRef::new(0, 0), note());
        if backing == "missing" {
            sheet.preserved = None;
        } else if backing != "kept" {
            let vml = Arc::make_mut(sheet.preserved.as_mut().unwrap()).vml.as_mut().unwrap();
            let text = std::str::from_utf8(&vml.data).unwrap();
            let text = if backing == "collision" {
                text.replace("_x0000_s1025", "_x0000_s1&#48;25")
            } else {
                text.replace("<xml ", "<v:xml ").replace("</xml>", "</v:xml>")
            };
            vml.data = Arc::from(text.into_bytes());
        }
        assert_eq!(sheet.shapes[0].info.compat_spid.as_deref(), Some("_x0000_s1025"));
        let bytes = write_xlsx(&wb).unwrap();
        let mut pkg = Package::open(&bytes).unwrap();
        let drawing = pkg.read_xml("xl/drawings/drawing1.xml").unwrap().unwrap();
        let link = drawing.children[0].path(&["sp", "nvSpPr", "cNvPr", "extLst", "ext", "compatExt"]);
        let vml = pkg.read_xml("xl/drawings/vmlDrawing1.vml").unwrap().unwrap();
        let note = vml.kids("shape").find(|s| s.child("ClientData").is_some_and(|c| c.attr("ObjectType") == Some("Note"))).unwrap();
        if backing == "kept" {
            let spid = link.unwrap().attr("spid").unwrap();
            assert_ne!(note.attr("id"), Some(spid));
            assert!(vml.kids("shape").any(|s| s.attr("id") == Some(spid) && s.child("ClientData").unwrap().attr("ObjectType") == Some("Button")));
        } else {
            assert!(link.is_none());
            assert_eq!(note.attr("id"), Some("_x0000_s1025"));
        }
    }
}

#[test]
fn compatibility_links_require_real_unique_non_note_backing_shapes() {
    for vml in [
        r#"<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shapetype id="_x0000_s1025"/></xml>"#,
        r#"<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="_x0000_s1026"/></xml>"#,
        r#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="_x0000_s1025"><x:ClientData ObjectType="Note"/></v:shape></xml>"#,
        r#"<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="_x0000_s1025"/><v:shape id="_x0000_s1025"/></xml>"#,
    ] {
        let (mut wb, _) = read_xlsx(&legacy_control(COMPAT, false)).unwrap();
        let sheet = wb.sheet_mut(0).unwrap();
        Arc::make_mut(sheet.preserved.as_mut().unwrap()).vml.as_mut().unwrap().data = Arc::from(vml.as_bytes());
        let bytes = write_xlsx(&wb).unwrap();
        let mut pkg = Package::open(&bytes).unwrap();
        let drawing = pkg.read_xml("xl/drawings/drawing1.xml").unwrap().unwrap();
        assert!(drawing.children[0].path(&["sp", "nvSpPr", "cNvPr", "extLst"]).is_none(), "{vml}");
        let bytes = replace_part(&legacy_control(COMPAT, false), "xl/drawings/vmlDrawing1.vml", Some(vml.as_bytes()));
        let (wb, report) = read_xlsx(&bytes).unwrap();
        assert!(wb.sheet(0).unwrap().shapes[0].info.compat_spid.is_none());
        assert!(report.warnings.iter().any(|w| w.contains("Legacy button") && w.contains("compatibility link removed")));
    }
}

#[test]
fn copied_objects_cannot_share_one_legacy_association() {
    let (mut wb, _) = read_xlsx(&legacy_control(COMPAT, false)).unwrap();
    let sheet = wb.sheet_mut(0).unwrap();
    let copy = Shape { id: 99, ..sheet.shapes[0].clone() };
    sheet.shapes.push(copy);
    let bytes = write_xlsx(&wb).unwrap();
    let mut pkg = Package::open(&bytes).unwrap();
    let drawing = pkg.read_xml("xl/drawings/drawing1.xml").unwrap().unwrap();
    let mut props = vec![];
    properties(&drawing, &mut props);
    assert_eq!(props.len(), 2);
    assert_eq!(props[0].path(&["extLst", "ext", "compatExt"]).unwrap().attr("spid"), Some("_x0000_s1025"));
    assert!(props[1].child("extLst").is_none());
}
