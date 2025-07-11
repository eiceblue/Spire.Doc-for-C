#include "pch.h"
using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/AddShapeGroup.docx";

	//create a document
	intrusive_ptr<Document> doc = new Document();
	intrusive_ptr<Section> sec = doc->AddSection();

	//add a new paragraph
	intrusive_ptr<Paragraph> para = sec->AddParagraph();
	//add a shape group with the height and width
	intrusive_ptr<ShapeGroup> shapegroup = para->AppendShapeGroup(375, 462);
	shapegroup->SetHorizontalPosition(180);
	//calcuate the scale ratio
	float X = static_cast<float>(shapegroup->GetWidth() / 1000.0f);
	float Y = static_cast<float>(shapegroup->GetHeight() / 1000.0f);

	intrusive_ptr<TextBox> txtBox = new TextBox(doc);
	txtBox->SetShapeType(ShapeType::RoundRectangle);
	txtBox->SetWidth(125 / X);
	txtBox->SetHeight(54 / Y);
	intrusive_ptr<Paragraph> paragraph = txtBox->GetBody()->AddParagraph();
	paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
	paragraph->AppendText(L"Start");
	txtBox->SetHorizontalPosition(19 / X);
	txtBox->SetVerticalPosition(27 / Y);
	txtBox->GetFormat()->SetLineColor(Color::GetGreen());
	shapegroup->GetChildObjects()->Add(txtBox);

	intrusive_ptr<ShapeObject> arrowLineShape = new ShapeObject(doc, ShapeType::DownArrow);
	arrowLineShape->SetWidth(16 / X);
	arrowLineShape->SetHeight(40 / Y);
	arrowLineShape->SetHorizontalPosition(69 / X);
	arrowLineShape->SetVerticalPosition(87 / Y);
	arrowLineShape->SetStrokeColor(Color::GetPurple());
	shapegroup->GetChildObjects()->Add(arrowLineShape);

	txtBox = new TextBox(doc);
	txtBox->SetShapeType(ShapeType::Rectangle);
	txtBox->SetWidth(125 / X);
	txtBox->SetHeight(54 / Y);
	paragraph = txtBox->GetBody()->AddParagraph();
	paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
	paragraph->AppendText(L"Step 1");
	txtBox->SetHorizontalPosition(19 / X);
	txtBox->SetVerticalPosition(131 / Y);
	txtBox->GetFormat()->SetLineColor(Color::GetBlue());
	shapegroup->GetChildObjects()->Add(txtBox);

	arrowLineShape = new ShapeObject(doc, ShapeType::DownArrow);
	arrowLineShape->SetWidth(16 / X);
	arrowLineShape->SetHeight(40 / Y);
	arrowLineShape->SetHorizontalPosition(69 / X);
	arrowLineShape->SetVerticalPosition(192 / Y);
	arrowLineShape->SetStrokeColor(Color::GetPurple());
	shapegroup->GetChildObjects()->Add(arrowLineShape);

	txtBox = new TextBox(doc);
	txtBox->SetShapeType(ShapeType::Parallelogram);
	txtBox->SetWidth(149 / X);
	txtBox->SetHeight(59 / Y);
	paragraph = txtBox->GetBody()->AddParagraph();
	paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
	paragraph->AppendText(L"Step 2");
	txtBox->SetHorizontalPosition(7 / X);
	txtBox->SetVerticalPosition(236 / Y);
	txtBox->GetFormat()->SetLineColor(Color::GetBlueViolet());
	shapegroup->GetChildObjects()->Add(txtBox);

	arrowLineShape = new ShapeObject(doc, ShapeType::DownArrow);
	arrowLineShape->SetWidth(16 / X);
	arrowLineShape->SetHeight(40 / Y);
	arrowLineShape->SetHorizontalPosition(66 / X);
	arrowLineShape->SetVerticalPosition(300 / Y);
	arrowLineShape->SetStrokeColor(Color::GetPurple());
	shapegroup->GetChildObjects()->Add(arrowLineShape);

	txtBox = new TextBox(doc);
	txtBox->SetShapeType(ShapeType::Rectangle);
	txtBox->SetWidth(125 / X);
	txtBox->SetHeight(54 / Y);
	paragraph = txtBox->GetBody()->AddParagraph();
	paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
	paragraph->AppendText(L"Step 3");
	txtBox->SetHorizontalPosition(19 / X);
	txtBox->SetVerticalPosition(345 / Y);
	txtBox->GetFormat()->SetLineColor(Color::GetBlue());
	shapegroup->GetChildObjects()->Add(txtBox);

	//save the document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2010);
	doc->Close();
}