#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/SetEditableRange.docx";
	std::wstring inputFile = DATAPATH"/SetEditableRange.docx";

	//Create a new document
	intrusive_ptr<Document> document = new Document();
	//Load file from disk
	document->LoadFromFile(inputFile.c_str());
	//Protect whole document
	document->Protect(ProtectionType::AllowOnlyReading, L"password");
	//Create tags for permission start and end
	intrusive_ptr<PermissionStart> start = new PermissionStart(document, L"testID");
	intrusive_ptr<PermissionEnd> end = new PermissionEnd(document, L"testID");
	//Add the start and end tags to allow the first paragraph to be edited.
	document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0)->GetChildObjects()->Insert(0, start);
	document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0)->GetChildObjects()->Add(end);
	//Save the document
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}