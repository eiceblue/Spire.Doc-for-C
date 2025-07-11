#include "pch.h"


using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring outputFile = output_path + L"HideParagraph.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Get the first section and the first paragraph from the word document.
	intrusive_ptr<Section> sec = document->GetSections()->GetItemInSectionCollection(0);
	intrusive_ptr<Paragraph> para = sec->GetParagraphs()->GetItemInParagraphCollection(0);

	//Loop through the textranges and set CharacterFormat.Hidden property as true to hide the texts.
	for (int i = 0; i < para->GetChildObjects()->GetCount(); i++)
	{
		intrusive_ptr<DocumentObject> obj = para->GetChildObjects()->GetItem(i);
		if (Object::CheckType<TextRange>(obj))
		{
			intrusive_ptr<TextRange> range = boost::dynamic_pointer_cast<TextRange>(obj);
			range->GetCharacterFormat()->SetHidden(true);
		}
	}

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}

