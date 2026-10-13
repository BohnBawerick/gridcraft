use std::sync::Arc;

use gridcraft_engine::Session;
use gridcraft_engine::core::CellRef;
use gridcraft_engine::model::{DrawingInfo, Image, Sheet};
use serde_json::json;

fn reopen(session: &mut Session) -> Session {
    let saved = session.execute("file.saveBytes", json!({"format": "xlsx"})).unwrap();
    let mut reopened = Session::new();
    reopened.execute("file.open", json!({"name": "drawings.xlsx", "base64": saved["base64"]})).unwrap();
    reopened
}

fn workbook() -> Session {
    let mut session = Session::new();
    session.new_workbook();
    session.execute("range.setValues", json!({"range": "A1", "values": [[1], [2], [3]]})).unwrap();
    session.execute("insert.chart", json!({"range": "A1:A3", "type": "line"})).unwrap();
    session.execute("insert.shape", json!({"kind": "rectangle", "at": "B2"})).unwrap();
    let sheet = Arc::make_mut(&mut session.doc_mut().unwrap().wb).sheet_mut(0).unwrap();
    let chart = sheet.charts[0].clone();
    let shape = sheet.shapes[0].clone();
    sheet.charts.clear();
    sheet.shapes.clear();
    for i in 0..3 {
        let info = |kind: &str, offset| DrawingInfo {
            name: Some(format!("{kind}{i}")),
            file_id: Some(40 + i * 3 + offset),
            order: Some(i * 3 + offset),
            ..Default::default()
        };
        let mut chart = chart.clone();
        chart.id = 100 + i;
        chart.info = info("chart", 1);
        sheet.charts.push(chart);
        let mut shape = shape.clone();
        shape.id = 200 + i;
        shape.info = info("shape", 0);
        sheet.shapes.push(shape.clone());
        sheet.images.push(Image {
            id: 300 + i,
            info: info("image", 2),
            anchor: shape.anchor,
            data: vec![1, 2, 3],
            mime: "image/png".into(),
            alt: String::new(),
        });
    }
    session
}

fn infos(sheet: &Sheet) -> Vec<(u32, &DrawingInfo)> {
    sheet
        .shapes
        .iter()
        .map(|o| (o.id, &o.info))
        .chain(sheet.images.iter().map(|o| (o.id, &o.info)))
        .chain(sheet.charts.iter().map(|o| (o.id, &o.info)))
        .collect()
}

fn stack(session: &Session, sheet: usize) -> Vec<(String, Option<u32>)> {
    let mut objects = infos(session.doc().unwrap().wb.sheet(sheet).unwrap());
    objects.sort_by_key(|(_, info)| info.order);
    objects.into_iter().map(|(_, info)| (info.name.clone().unwrap(), info.file_id)).collect()
}

#[test]
fn drawing_identity_arrange_survives_save_undo_and_redo_for_every_kind_and_direction() {
    for kind in ["shape", "chart", "image"] {
        for (command, extreme, from, to) in [
            ("arrange.bringForward", false, 0, 1),
            ("arrange.sendBackward", false, 2, 1),
            ("arrange.bringForward", true, 0, 2),
            ("arrange.sendBackward", true, 2, 0),
            ("arrange.bringForward", false, 2, 2),
            ("arrange.sendBackward", false, 0, 0),
        ] {
            let mut session = reopen(&mut workbook());
            let before = stack(&session, 0);
            let target = format!("{kind}{from}");
            let id =
                infos(session.doc().unwrap().wb.sheet(0).unwrap()).into_iter().find(|(_, info)| info.name.as_deref() == Some(&target)).unwrap().0;
            let slots: Vec<_> = before.iter().enumerate().filter(|(_, o)| o.0.starts_with(kind)).map(|(i, _)| i).collect();
            let mut reordered: Vec<_> = slots.iter().map(|i| before[*i].clone()).collect();
            let item = reordered.remove(from);
            reordered.insert(to, item);
            let mut expected = before.clone();
            for (slot, item) in slots.iter().zip(reordered) {
                expected[*slot] = item;
            }
            let mut params = json!({"kind": kind, "id": id});
            params[if command == "arrange.bringForward" { "toFront" } else { "toBack" }] = json!(extreme);
            session.execute(command, params).unwrap();
            assert_eq!(stack(&reopen(&mut session), 0), expected, "{kind} {command} {extreme}");
            session.execute("edit.undo", json!({})).unwrap();
            assert_eq!(stack(&reopen(&mut session), 0), before);
            session.execute("edit.redo", json!({})).unwrap();
            assert_eq!(stack(&reopen(&mut session), 0), expected);
        }
    }
}

#[test]
fn drawing_identity_arrange_assigns_positions_to_new_objects() {
    let mut session = workbook();
    let sheet = Arc::make_mut(&mut session.doc_mut().unwrap().wb).sheet_mut(0).unwrap();
    for shape in &mut sheet.shapes {
        shape.info.order = None;
    }
    session.execute("arrange.bringForward", json!({"kind": "shape", "id": 200, "toFront": true})).unwrap();
    let back = reopen(&mut session);
    let names: Vec<_> = back.doc().unwrap().wb.sheet(0).unwrap().shapes.iter().map(|o| o.info.name.as_deref().unwrap()).collect();
    assert_eq!(names, ["shape1", "shape2", "shape0"]);
}

#[test]
fn drawing_identity_chart_transfer_preserves_destination_ids_and_undo() {
    let mut session = workbook();
    let wb = Arc::make_mut(&mut session.doc_mut().unwrap().wb);
    let source = wb.sheet_mut(0).unwrap();
    source.charts[0].info.file_id = Some(90);
    source.charts[0].info.order = Some(2);
    let mut shape = source.shapes[0].clone();
    shape.info.file_id = Some(90);
    shape.info.order = Some(2);
    let mut target = Sheet::new("Target");
    target.shapes.push(shape);
    wb.sheets.push(Arc::new(target));
    let mut session = reopen(&mut session);
    let before_source = stack(&session, 0);
    let before_target = stack(&session, 1);
    let chart = session.doc().unwrap().wb.sheet(0).unwrap().charts[0].clone();
    session.execute("chart.move", json!({"chart": chart.id, "sheet": "Target", "at": "C7"})).unwrap();
    let moved = &session.doc().unwrap().wb.sheet(1).unwrap().charts[0];
    assert_eq!(moved.info.file_id, None);
    assert_eq!(moved.info.order, None);
    let back = reopen(&mut session);
    let target = back.doc().unwrap().wb.sheet(1).unwrap();
    assert_eq!(target.shapes[0].info.file_id, Some(90));
    assert_ne!(target.charts[0].info.file_id, Some(90));
    assert!(target.charts[0].info.order > target.shapes[0].info.order);
    assert_eq!(target.charts[0].anchor.cell, CellRef::parse("C7").unwrap());
    session.execute("edit.undo", json!({})).unwrap();
    let back = reopen(&mut session);
    assert_eq!(stack(&back, 0), before_source);
    assert_eq!(stack(&back, 1), before_target);
    session.execute("edit.redo", json!({})).unwrap();
    assert_eq!(reopen(&mut session).doc().unwrap().wb.sheet(1).unwrap().shapes[0].info.file_id, Some(90));
}

#[test]
fn drawing_identity_chart_move_within_sheet_keeps_identity_and_order() {
    for explicit_sheet in [false, true] {
        let mut session = reopen(&mut workbook());
        let before = stack(&session, 0);
        let sheet = session.doc().unwrap().wb.sheet(0).unwrap();
        let chart = sheet.charts[0].clone();
        let chart_ids: Vec<_> = sheet.charts.iter().map(|o| o.id).collect();
        let mut params = json!({"chart": chart.id, "at": "C7"});
        if explicit_sheet {
            params["sheet"] = json!(sheet.name);
        }
        session.execute("chart.move", params).unwrap();
        let sheet = session.doc().unwrap().wb.sheet(0).unwrap();
        assert_eq!(sheet.charts.iter().map(|o| o.id).collect::<Vec<_>>(), chart_ids);
        assert_eq!(sheet.charts[0].info, chart.info);
        assert_eq!(sheet.charts[0].anchor.cell, CellRef::parse("C7").unwrap());
        assert_eq!(stack(&reopen(&mut session), 0), before);
    }
}

#[test]
fn drawing_identity_hidden_objects_do_not_change_print_bounds_pages_or_scale() {
    for kind in ["chart", "image", "shape"] {
        let mut session = workbook();
        let sheet = Arc::make_mut(&mut session.doc_mut().unwrap().wb).sheet_mut(0).unwrap();
        sheet.charts.truncate(usize::from(kind == "chart"));
        sheet.images.truncate(usize::from(kind == "image"));
        sheet.shapes.truncate(usize::from(kind == "shape"));
        for (anchor, info) in sheet
            .charts
            .iter_mut()
            .map(|o| (&mut o.anchor, &mut o.info))
            .chain(sheet.images.iter_mut().map(|o| (&mut o.anchor, &mut o.info)))
            .chain(sheet.shapes.iter_mut().map(|o| (&mut o.anchor, &mut o.info)))
        {
            anchor.cell = CellRef::parse("Z1000").unwrap();
            info.hidden = true;
        }
        for fit in [false, true] {
            let preview = session.execute("file.printPreview", json!({"fitToPage": fit})).unwrap();
            assert_eq!(preview["pages"], 1, "{kind}");
            assert_eq!(preview["pageList"][0]["range"], "A1:A3", "{kind}");
            assert_eq!(preview["pageList"][0]["scale"], 100.0, "{kind}");
            let pdf = session.execute("file.exportPdf", json!({"fitToPage": fit})).unwrap();
            assert_eq!(pdf["pages"], 1, "{kind}");
        }
        let sheet = Arc::make_mut(&mut session.doc_mut().unwrap().wb).sheet_mut(0).unwrap();
        for info in sheet
            .charts
            .iter_mut()
            .map(|o| &mut o.info)
            .chain(sheet.images.iter_mut().map(|o| &mut o.info))
            .chain(sheet.shapes.iter_mut().map(|o| &mut o.info))
        {
            info.hidden = false;
        }
        let preview = session.execute("file.printPreview", json!({"fitToPage": false})).unwrap();
        assert!(preview["pages"].as_u64().unwrap() > 1, "{kind}");
    }
}
