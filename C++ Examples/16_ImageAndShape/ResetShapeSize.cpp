#include "pch.h"
using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/ResetShapeSize.docx";
	std::wstring inputFile = DATAPATH"/Shapes.docx";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section and the first paragraph that contains the shape
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);
	intrusive_ptr<Paragraph> para = section->GetParagraphs()->GetItemInParagraphCollection(0);

	//Get the second shape and reset the width and height for the shape
	intrusive_ptr<ShapeObject> shape = Object::Dynamic_cast<ShapeObject>(para->GetChildObjects()->GetItem(1));
	shape->SetWidth(200);
	shape->SetHeight(200);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();

}