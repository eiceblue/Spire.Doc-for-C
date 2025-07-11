#include "pch.h"
using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/RemoveShape.docx";
	std::wstring inputFile = DATAPATH"/Shapes.docx";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	//Get all the child objects of paragraph
	for (int i = 0; i < section->GetParagraphs()->GetCount(); i++)
	{
		intrusive_ptr<Paragraph> para = section->GetParagraphs()->GetItemInParagraphCollection(i);
		for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
		{
			//If the child objects is shape object
			if (Object::CheckType<ShapeObject>(para->GetChildObjects()->GetItem(j)))
			{
				//Remove the shape object
				para->GetChildObjects()->RemoveAt(j);
				--j;
			}
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}