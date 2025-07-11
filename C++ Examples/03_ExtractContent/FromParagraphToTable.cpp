#include "pch.h"

using namespace Spire::Doc;

void ExtractByTable(intrusive_ptr<Document> sourceDocument, intrusive_ptr<Document> destinationDocument, int startPara, int tableNo)
{
	//Get the table from the source document
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetTables()->GetItemInTableCollection(tableNo - 1));

	//Get the table index
	int index = sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->IndexOf(table);
	for (int i = startPara - 1; i <= index; i++)
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
	wstring inputFile = input_path + L"IncludingTable.docx";
	wstring outputFile = output_path + L"FromParagraphToTable.docx";

	//Create the first document
	intrusive_ptr<Document> sourceDocument = new Document();

	//Load the source document from disk.
	sourceDocument->LoadFromFile(inputFile.c_str());

	//Create a destination document
	intrusive_ptr<Document> destinationDoc = new Document();

	//Add a section
	intrusive_ptr<Section> destinationSection = destinationDoc->AddSection();

	//Extract the content from the first paragraph to the first table
	ExtractByTable(sourceDocument, destinationDoc, 1, 1);

	//Save the document.
	destinationDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	destinationDoc->Close();
}


