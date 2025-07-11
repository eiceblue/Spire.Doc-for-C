#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/LockHeader.docx";
	std::wstring inputFile = DATAPATH"/HeaderAndFooter.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	//Protect the document and set the ProtectionType as AllowOnlyFormFields
	doc->Protect(ProtectionType::AllowOnlyFormFields, L"123");

	//Set the ProtectForm as false to unprotect the section
	section->SetProtectForm(false);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}