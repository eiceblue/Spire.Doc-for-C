#include "pch.h"


using namespace Spire::Doc;


int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Text1.docx";
	wstring outputFile = output_path + L"ChangeCase.docx";

	// Create a new document and load from file;
	intrusive_ptr<Document> doc =  new Document();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<TextRange> textRange;
	//Get the first paragraph and set its CharacterFormat to AllCaps
	intrusive_ptr<Paragraph> para1 = doc->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(1);

	for (int i = 0; i < para1->GetChildObjects()->GetCount(); i++)
	{
		intrusive_ptr<DocumentObject> obj = para1->GetChildObjects()->GetItem(i);
		if (Object::CheckType<TextRange>(obj))
		{
			textRange = boost::dynamic_pointer_cast<TextRange>(obj);
			textRange->GetCharacterFormat()->SetAllCaps(true);
		}
	}

	//Get the third paragraph and set its CharacterFormat to IsSmallCaps
	intrusive_ptr<Paragraph> para2 = doc->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(3);
	for (int i = 0; i < para2->GetChildObjects()->GetCount(); i++)
	{
		intrusive_ptr<DocumentObject> obj = para2->GetChildObjects()->GetItem(i);
		if (Object::CheckType<TextRange>(obj))
		{
			textRange = boost::dynamic_pointer_cast<TextRange>(obj);
			textRange->GetCharacterFormat()->SetIsSmallCaps(true);
		}
	}

	//Save and launch the document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();

}