use egui::accesskit::Role;
use egui_kittest::kittest::Queryable;
use gridcraft_engine::{Session, core::CellRef, model::FontScheme};
use gridcraft_ui_egui::{SheetApp, dialogs};
use serde_json::json;

fn harness(session: Session) -> egui_kittest::Harness<'static, SheetApp> {
    let mut app = SheetApp::new(session, Default::default());
    app.open_dialog("formatCells", json!({"tab": "Font"}));
    let mut h = egui_kittest::Harness::builder().with_size(egui::vec2(1200.0, 800.0)).build_ui_state(|ui, app| dialogs::show(app, ui.ctx()), app);
    h.run_steps(6);
    h
}

fn confirm_and_check(h: &mut egui_kittest::Harness<'static, SheetApp>, cell: &str, face: &str, scheme: FontScheme) {
    h.get_by_label("   OK   ").click();
    h.run_steps(6);
    assert!(h.state().dialog.is_none());
    let wb = &h.state().session.active().unwrap().wb;
    let at = CellRef::parse(cell).unwrap();
    let font = &wb.styles.get(wb.active().unwrap().style_id(at)).font;
    assert_eq!((font.name.as_str(), font.scheme), (face, scheme));
    let (back, _) = gridcraft_xlsx::read_xlsx(&gridcraft_xlsx::write_xlsx(wb).unwrap()).unwrap();
    let font = &back.styles.get(back.active().unwrap().style_id(at)).font;
    assert_eq!((font.name.as_str(), font.scheme), (face, scheme));
}

#[test]
fn font_scheme_dialog_font_choices_are_explicit() {
    for choices in [vec!["Calibri"], vec!["Arial", "Calibri"]] {
        let mut session = Session::new();
        session.new_workbook();
        let mut h = harness(session);
        for face in choices {
            h.get_by_role(Role::ComboBox).click();
            h.run_steps(6);
            h.get_by_label(face).click();
            h.run_steps(6);
            assert_eq!(h.state().dialog.as_ref().unwrap().values["style"]["font"]["scheme"], "None");
        }
        confirm_and_check(&mut h, "A1", "Calibri", FontScheme::None);
    }
}

#[test]
fn font_scheme_dialog_nonfont_edits_keep_theme_links() {
    for (style, face, scheme) in [("Good", "Arial", FontScheme::Minor), ("Title", "Cambria", FontScheme::Major)] {
        let mut session = Session::new();
        session.new_workbook();
        session.execute("pageLayout.theme", json!({"minorFont": "Arial", "majorFont": "Cambria"})).unwrap();
        session.execute("home.cellStyle", json!({"range": "A1", "name": style})).unwrap();
        let mut h = harness(session);
        h.get_by_label("Bold").click();
        h.run_steps(6);
        confirm_and_check(&mut h, "A1", face, scheme);
    }
}

#[test]
fn font_scheme_dialog_full_explicit_style_clears_other_cells_theme_links() {
    let mut session = Session::new();
    session.new_workbook();
    session.execute("home.fontName", json!({"range": "A1", "name": "Calibri"})).unwrap();
    session.execute("selection.set", json!({"range": "A1:B1"})).unwrap();
    let mut h = harness(session);
    assert_eq!(h.state().dialog.as_ref().unwrap().values["style"]["font"]["scheme"], "None");
    confirm_and_check(&mut h, "B1", "Calibri", FontScheme::None);
}
