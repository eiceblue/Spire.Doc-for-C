#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Footnote.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveFootnote.docx";

	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

	//Traverse paragraphs in the section and find the footnote
	for (int k = 0; k < section->GetParagraphs()->GetCount(); k++)
	{
		intrusive_ptr<Paragraph> para = section->GetParagraphs()->GetItemInParagraphCollection(k);
		int index = -1;
		for (int i = 0, cnt = para->GetChildObjects()->GetCount(); i < cnt; i++)
		{
			intrusive_ptr<ParagraphBase> pBase = Object::Dynamic_cast<ParagraphBase>(para->GetChildObjects()->GetItem(i));
			if (Object::CheckType<Footnote>(pBase))
			{
				index = i;
				break;
			}
		}

		if (index > -1)
		{
			//remove the footnote
			para->GetChildObjects()->RemoveAt(index);
		}
	}

	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
