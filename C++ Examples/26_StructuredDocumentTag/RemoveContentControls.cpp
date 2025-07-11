#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"RemoveContentControls.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveContentControls.docx";

	//Load document from disk
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Loop through sections
	for (int s = 0; s < doc->GetSections()->GetCount(); s++)
	{
		intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(s);
		for (int i = 0; i < section->GetBody()->GetChildObjects()->GetCount(); i++)
		{
			//Loop through contents in paragraph
			if (Object::CheckType<Paragraph>(section->GetBody()->GetChildObjects()->GetItem(i)) )
			{
				intrusive_ptr<Paragraph> para = boost::dynamic_pointer_cast<Paragraph>(section->GetBody()->GetChildObjects()->GetItem(i));
				for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
				{
					//Find the StructureDocumentTagInline
					if (Object::CheckType<StructureDocumentTagInline>(para->GetChildObjects()->GetItem(j)))
					{
						intrusive_ptr<StructureDocumentTagInline> sdt = boost::dynamic_pointer_cast<StructureDocumentTagInline>(para->GetChildObjects()->GetItem(j));
						//Remove the content control from paragraph
						para->GetChildObjects()->Remove(sdt);
						j--;
					}
				}
			}
			if (Object::CheckType<StructureDocumentTag>(section->GetBody()->GetChildObjects()->GetItem(i)))
			{
				intrusive_ptr<StructureDocumentTag> sdt = boost::dynamic_pointer_cast<StructureDocumentTag>(section->GetBody()->GetChildObjects()->GetItem(i));
				section->GetBody()->GetChildObjects()->Remove(sdt);
				i--;
			}
		}
	}

	//Save the Word document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
}
