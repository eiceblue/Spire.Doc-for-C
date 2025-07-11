#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Toc.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ChangeTOCStyle.docx";

	//Load document from disk
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Defind a Toc style
	intrusive_ptr<ParagraphStyle> tocStyle = Object::Dynamic_cast<ParagraphStyle>(Style::CreateBuiltinStyle(BuiltinStyle::Toc1, doc));
	tocStyle->GetCharacterFormat()->SetFontName(L"Aleo");
	tocStyle->GetCharacterFormat()->SetFontSize(15.0f);
	tocStyle->GetCharacterFormat()->SetTextColor(Color::GetCadetBlue());
	doc->GetStyles()->Add(tocStyle);

	//Loop through sections
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(i);
		//Loop through content of section
		for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++)
		{
			intrusive_ptr<DocumentObject> obj = section->GetBody()->GetChildObjects()->GetItem(j);
			//Find the structure document tag
			if (Object::CheckType<StructureDocumentTag>(obj))
			{
				intrusive_ptr<StructureDocumentTag> tag = boost::dynamic_pointer_cast<StructureDocumentTag>(obj);
				//Find the paragraph where the TOC1 locates
				for (int k = 0; k < tag->GetChildObjects()->GetCount(); k++)
				{
					intrusive_ptr<DocumentObject> cObj = tag->GetChildObjects()->GetItem(k);
					if (Object::CheckType<Paragraph>(cObj))
					{
						intrusive_ptr<Paragraph> para = boost::dynamic_pointer_cast<Paragraph>(cObj);
						if (wcscmp(para->GetStyleName(), L"TOC1") == 0)
						{
							//Apply the new style for TOC1 paragraph
							para->ApplyStyle(tocStyle->GetName());
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