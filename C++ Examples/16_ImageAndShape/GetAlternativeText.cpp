#include "pch.h"
using namespace Spire::Doc;

int main() {
	std::wstring inputFile = DATAPATH"/ShapeWithAlternativeText.docx";
	std::wstring outputFile = OUTPUTPATH"/GetAlternativeText.txt";


	//Create a document
	intrusive_ptr<Document> document = new Document();
	//Create string builder
	std::wstring builder;
	document->LoadFromFile(inputFile.c_str());

	//Loop through shapes and get the AlternativeText
	for (int i = 0; i < document->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
		for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
		{
			intrusive_ptr<Paragraph> para = section->GetParagraphs()->GetItemInParagraphCollection(j);
			for (int k = 0; k < para->GetChildObjects()->GetCount(); k++)
			{
				intrusive_ptr<DocumentObject> obj = para->GetChildObjects()->GetItem(k);
				if (Object::CheckType<ShapeObject>(obj))
				{
					std::wstring text = (boost::dynamic_pointer_cast<ShapeObject>(obj))->GetAlternativeText();
					//Append the alternative text in builder
					builder.append(text);
					builder.append(L"\n");
				}
			}
		}
	}

	//Save doc file.
	std::wofstream foo(outputFile);
	foo << builder.c_str();
	foo.close();
	document->Close();
}