//! Theme-following fonts and the hidden `_FilterDatabase` names survive save.

use gridcraft_core::{CellRef, RangeRef};
use gridcraft_model::{AutoFilter, FontScheme};

use super::minimal;
use crate::package::Package;
use crate::{read_xlsx, write_xlsx};

fn at(a: &str) -> CellRef {
    CellRef::parse(a).unwrap()
}

const THEME: &str = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements><a:clrScheme name="x"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="f"><a:majorFont><a:latin typeface="Heading Face"/></a:majorFont><a:minorFont><a:latin typeface="Body Face"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#;

/// Normal in the theme's body font; a font that follows the body font under another name; an
/// explicit font that only shares the body font's name; a font that follows the heading font.
const STYLES: &str = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="4"><font><sz val="11"/><name val="Body Face"/><scheme val="minor"/></font><font><sz val="11"/><name val="Old Face"/><scheme val="minor"/></font><font><b/><sz val="11"/><name val="Body Face"/></font><font><sz val="15"/><name val="Heading Face"/><scheme val="major"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders>
  <cellXfs count="4"><xf fontId="0"/><xf fontId="1"/><xf fontId="2"/><xf fontId="3"/></cellXfs></styleSheet>"#;

fn book() -> Vec<u8> {
    let rels = r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>"#;
    let body = r#"<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>2</v></c><c r="C1" s="3"><v>3</v></c></row></sheetData>"#;
    minimal(body, &[("xl/styles.xml", STYLES), ("xl/theme/theme1.xml", THEME)], rels, "")
}

/// (name, scheme) of every `<font>` in the saved styles part.
fn saved_fonts(bytes: &[u8]) -> Vec<(String, Option<String>)> {
    let mut pkg = Package::open(bytes).unwrap();
    let x = pkg.read_xml("xl/styles.xml").unwrap().unwrap();
    x.child("fonts")
        .unwrap()
        .kids("font")
        .map(|f| {
            (f.child("name").and_then(|n| n.attr("val")).unwrap_or("").to_string(), f.child("scheme").and_then(|s| s.attr("val")).map(str::to_string))
        })
        .collect()
}

#[test]
fn theme_fonts_stay_theme_fonts_and_explicit_fonts_stay_explicit() {
    let (wb, _) = read_xlsx(&book()).unwrap();
    let s = wb.sheet(0).unwrap();
    let font = |a: &str| wb.styles.get(s.cell(at(a)).unwrap().style).font.clone();
    // Excel shows the theme's body font for a font that follows it, whatever name it carries.
    assert_eq!((font("A1").name.as_str(), font("A1").scheme), ("Body Face", FontScheme::Minor));
    assert_eq!((font("B1").name.as_str(), font("B1").scheme), ("Body Face", FontScheme::None));
    assert_eq!((font("C1").name.as_str(), font("C1").scheme), ("Heading Face", FontScheme::Major));

    let saved = write_xlsx(&wb).unwrap();
    let fonts = saved_fonts(&saved);
    let has = |name: &str, scheme: Option<&str>| fonts.iter().any(|(n, s)| n == name && s.as_deref() == scheme);
    assert!(has("Body Face", Some("minor")), "{fonts:?}");
    assert!(has("Body Face", None), "the explicit font does not become a theme font: {fonts:?}");
    assert!(has("Heading Face", Some("major")), "{fonts:?}");

    let (back, _) = read_xlsx(&saved).unwrap();
    let b = back.sheet(0).unwrap();
    for a in ["A1", "B1", "C1"] {
        assert_eq!(back.styles.get(b.cell(at(a)).unwrap().style).font, font(a), "{a}");
    }
}

#[test]
fn filter_database_names_survive_save() {
    let names = r#"<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$B$9</definedName></definedNames>"#;
    // The workbook XML takes the names after <sheets>: put them in through the sheet list.
    let bytes = minimal("<sheetData/>", &[], "", "");
    let mut z = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut parts: Vec<(String, Vec<u8>)> = vec![];
    for i in 0..z.len() {
        let mut f = z.by_index(i).unwrap();
        let mut d = vec![];
        std::io::Read::read_to_end(&mut f, &mut d).unwrap();
        if f.name() == "xl/workbook.xml" {
            d = String::from_utf8(d).unwrap().replace("</sheets>", &format!("</sheets>{names}")).into_bytes();
        }
        parts.push((f.name().to_string(), d));
    }
    let refs: Vec<(&str, &[u8])> = parts.iter().map(|(n, d)| (n.as_str(), d.as_slice())).collect();
    let (mut wb, _) = read_xlsx(&super::make_zip(&refs)).unwrap();
    let kept = wb.names.iter().find(|n| n.name == "_xlnm._FilterDatabase").expect("kept");
    assert!(kept.hidden && kept.scope == Some(0), "{kept:?}");

    // A sheet without an AutoFilter keeps the name it came with.
    let filter_names = |bytes: &[u8]| -> Vec<(Option<String>, bool, String)> {
        let mut pkg = Package::open(bytes).unwrap();
        let x = pkg.read_xml("xl/workbook.xml").unwrap().unwrap();
        x.child("definedNames")
            .map(|d| {
                d.kids("definedName")
                    .filter(|n| n.attr("name") == Some("_xlnm._FilterDatabase"))
                    .map(|n| (n.attr("localSheetId").map(str::to_string), n.flag("hidden", false), n.text.clone()))
                    .collect()
            })
            .unwrap_or_default()
    };
    assert_eq!(filter_names(&write_xlsx(&wb).unwrap()), [(Some("0".into()), true, "Data!$A$1:$B$9".into())]);

    // With an AutoFilter the name is the filter's range, written once.
    wb.sheet_mut(0).unwrap().autofilter = Some(AutoFilter { range: RangeRef::parse("A1:C20").unwrap(), criteria: vec![] });
    let saved = write_xlsx(&wb).unwrap();
    assert_eq!(filter_names(&saved), [(Some("0".into()), true, "Data!$A$1:$C$20".into())]);
    // Reopened, the AutoFilter carries the range again and writes the same name.
    let (back, _) = read_xlsx(&saved).unwrap();
    assert!(back.names.iter().all(|n| n.name != "_xlnm._FilterDatabase"));
    assert_eq!(back.sheet(0).unwrap().autofilter.as_ref().map(|af| af.range), RangeRef::parse("A1:C20"));
    assert_eq!(filter_names(&write_xlsx(&back).unwrap()), filter_names(&saved));
}
