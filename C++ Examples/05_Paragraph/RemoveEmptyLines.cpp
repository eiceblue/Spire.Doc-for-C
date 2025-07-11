#include "pch.h"


using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Docx_3.docx";
	wstring outputFile = output_path + L"RemoveEmptyLines.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Traverse every section on the word document and remove the null and empty paragraphs.
	for (int i = 0; i < document->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
		//int secChildObjectsCount = ;
		for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++)
		{
			intrusive_ptr<DocumentObject> secChildObject = section->GetBody()->GetChildObjects()->GetItem(j);
			if (secChildObject->GetDocumentObjectType() == DocumentObjectType::Paragraph)
			{
				intrusive_ptr<Paragraph> para = Object::Dynamic_cast<Paragraph>(secChildObject);
				std::wstring paraText = para->GetText();
				if (paraText.empty())
				{
					section->GetBody()->GetChildObjects()->Remove(secChildObject);
					j--;
				}
			}

		}
	}

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();

}
