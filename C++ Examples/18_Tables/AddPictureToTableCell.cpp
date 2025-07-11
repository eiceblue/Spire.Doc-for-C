#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/AddPictureToTableCell.docx";
	std::wstring inputFile = DATAPATH"/TableTemplate.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first table from the first section of the document
	intrusive_ptr<Table> table1 = Object::Dynamic_cast<Table>(doc->GetSections()->GetItemInSectionCollection(0)->GetTables()->GetItemInTableCollection(0));

	//Add a picture to the specified table cell and set picture size
	intrusive_ptr<DocPicture> picture = table1->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(2)->GetParagraphs()->GetItemInParagraphCollection(0)->AppendPicture(DATAPATH"/Spire.Doc.png");
	
	picture->SetWidth(100);
	picture->SetHeight(100);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
