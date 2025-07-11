#include "pch.h"
using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/RotateShape.docx";
	std::wstring inputFile = DATAPATH"/Shapes.docx";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	//Traverse the word document and set the shape rotation as 20
	for (int i = 0; i < section->GetParagraphs()->GetCount(); i++)
	{
		intrusive_ptr<Paragraph> para = section->GetParagraphs()->GetItemInParagraphCollection(i);
		for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
		{
			intrusive_ptr<DocumentObject> obj = para->GetChildObjects()->GetItem(j);
			if (Object::CheckType<ShapeObject>(obj))
			{
				(boost::dynamic_pointer_cast<ShapeObject>(obj))->SetRotation(20.0);
			}
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();

}