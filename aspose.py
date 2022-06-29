import aspose.words as aw
# For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-Python-via-.NET
doc = aw.Document()
builder = aw.DocumentBuilder(doc)

shape = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, aw.drawing.RelativeHorizontalPosition.PAGE, 100,
    aw.drawing.RelativeVerticalPosition.PAGE, 100, 50, 50, aw.drawing.WrapType.NONE)
shape.rotation = 30.0

builder.writeln()

shape = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 50, 50)
shape.rotation = 30.0

saveOptions = aw.saving.OoxmlSaveOptions(aw.SaveFormat.DOCX)
saveOptions.compliance = aw.saving.OoxmlCompliance.ISO29500_2008_TRANSITIONAL



doc.save(docs_base.artifacts_dir + "WorkingWithShapes.insert_shape.docx", saveOptions)