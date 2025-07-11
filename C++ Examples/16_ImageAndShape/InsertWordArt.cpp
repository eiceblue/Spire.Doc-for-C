#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/InsertWordArt.docx";
	std::wstring inputFile = DATAPATH"/InsertWordArt.docx";

	//Create Word document.
	intrusive_ptr<Document> doc = new Document();

	//Load Word document.
	doc->LoadFromFile(inputFile.c_str());

	//Add a paragraph.
	intrusive_ptr<Paragraph> paragraph = doc->GetSections()->GetItemInSectionCollection(0)->AddParagraph();

	//Add a shape.
	intrusive_ptr<ShapeObject> shape = paragraph->AppendShape(250, 70, ShapeType::TextWave4);

	//Set the position of the shape.
	shape->SetVerticalPosition(20);
	shape->SetHorizontalPosition(80);

	//set the text of WordArt.
	shape->GetWordArt()->SetText(L"Thanks for reading.");

	//Set the fill color.
	shape->SetFillColor(Color::GetRed());

	//Set the border color of the text.
	shape->SetStrokeColor(Color::GetYellow());

	//Save docx file.
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
}