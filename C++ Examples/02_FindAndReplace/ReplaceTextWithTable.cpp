#include "pch.h"

using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring outputFile = output_path + L"ReplaceTextWithTable.docx";
	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Return TextSection by finding the key text string "Christmas Day, December 25".
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);
	intrusive_ptr<TextSelection> selection = document->FindString(L"Christmas Day, December 25", true, true);

	//Return TextRange from TextSection, then get OwnerParagraph through TextRange.
	intrusive_ptr<TextRange> range = selection->GetAsOneRange();
	intrusive_ptr<Paragraph> paragraph = range->GetOwnerParagraph();

	//Return the zero-based index of the specified paragraph.
	intrusive_ptr<Body> body = paragraph->GetOwnerTextBody();
	int index = body->GetChildObjects()->IndexOf(paragraph);

	//Create a new table.
	intrusive_ptr<Table> table = section->AddTable(true);
	table->ResetCells(3, 3);

	//Remove the paragraph and insert table into the collection at the specified index.
	body->GetChildObjects()->Remove(paragraph);
	body->GetChildObjects()->Insert(index, table);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}
