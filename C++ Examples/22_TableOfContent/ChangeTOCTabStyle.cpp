#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Toc.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ChangeTOCTabStyle.docx";

	//Load document from disk
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	////Loop through sections
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(i);
		//Loop through content of section
		for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++)
		{
			intrusive_ptr<DocumentObject> obj = section->GetBody()->GetChildObjects()->GetItem(j);
			//Find the structure document tag
			if (Object::CheckType<StructureDocumentTag>(obj) )
			{
				intrusive_ptr<StructureDocumentTag> tag = boost::dynamic_pointer_cast<StructureDocumentTag>(obj);
				//Find the paragraph where the TOC1 locates
				for (int k = 0; k < tag->GetChildObjects()->GetCount(); k++)
				{
					intrusive_ptr<DocumentObject> cObj = tag->GetChildObjects()->GetItem(k);
					if (Object::CheckType<Paragraph>(cObj) )
					{
						intrusive_ptr<Paragraph> para = boost::dynamic_pointer_cast<Paragraph>(cObj);
						if (wcscmp(para->GetStyleName(), L"TOC2") == 0)
						{
							//Set the tab style of paragraph
							for (int a = 0; a < para->GetFormat()->GetTabs()->GetCount(); a++)
							{
								intrusive_ptr<Tab> tab = para->GetFormat()->GetTabs()->GetItem(a);
								tab->SetPosition(tab->GetPosition() + 20);
								tab->SetTabLeader(TabLeader::NoLeader);
							}
						}
					}
				}
			}
		}
	}

	//Save the Word file
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
}
