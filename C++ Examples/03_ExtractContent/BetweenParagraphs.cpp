#include "pch.h"

using namespace Spire::Doc;

void ExtractBetweenParagraphs(intrusive_ptr<Document> sourceDocument, intrusive_ptr<Document> destinationDocument, int startPara, int endPara)
{
	//Extract the content
	for (int i = startPara - 1; i < endPara; i++)
	{
		//Clone the ChildObjects of source document
		intrusive_ptr<DocumentObject> doobj = sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->GetItem(i)->Clone();

		//Add to destination document 
		destinationDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->Add(doobj);
	}
}

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Sample.docx";
	wstring outputFile = output_path + L"BetweenParagraphs.docx";

	//Create the first document
	intrusive_ptr<Document> sourceDocument = new Document();

	//Load the source document from disk.
	sourceDocument->LoadFromFile(inputFile.c_str());

	//Create a destination document
	intrusive_ptr<Document> destinationDoc = new Document();

	//Add a section
	intrusive_ptr<Section> section = destinationDoc->AddSection();

	//Extract content between the first paragraph to the third paragraph
	ExtractBetweenParagraphs(sourceDocument, destinationDoc, 1, 3);

	//Save the document.
	destinationDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	destinationDoc->Close();

}
