#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile_1 = input_path + L"Insert.docx";
	wstring inputFile_2 = input_path + L"TableOfContent.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"MergeDocsOnSamePage.docx";
	//Create a document
	intrusive_ptr<Document> document = new Document();


	//Load the source document from disk.
	document->LoadFromFile(inputFile_1.c_str());

	//Clone a destination  document
	intrusive_ptr<Document> destinationDocument = new Document();


	//Load the destination document from disk.
	destinationDocument->LoadFromFile(inputFile_2.c_str());

	//Traverse sections
	int sectionCount = document->GetSections()->GetCount();
	for (int i = 0; i < sectionCount; i++)
	{
		intrusive_ptr<Section> sec = document->GetSections()->GetItemInSectionCollection(i);
		int sectionChildObjectsCount = sec->GetBody()->GetChildObjects()->GetCount();
		for (int j = 0; j < sectionChildObjectsCount; j++)
		{
			intrusive_ptr<DocumentObject> documentObject = sec->GetBody()->GetChildObjects()->GetItem(j);
			destinationDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->Add(documentObject->Clone());
		}
	}
	//Save the document.
	destinationDocument->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	destinationDocument->Close();

}
