#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/RemoveCustomPropertyFields.docx";
	std::wstring inputFile = DATAPATH"/RemoveCustomPropertyFields.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Get custom document properties object.
	intrusive_ptr<CustomDocumentProperties> cdp = document->GetCustomDocumentProperties();

	//Remove all custom property fields in the document.
	for (int i = 0; i < cdp->GetCount();/* i++*/)
	{
		cdp->Remove(cdp->GetItem(i)->GetName());
	}

	document->SetIsUpdateFields(true);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}