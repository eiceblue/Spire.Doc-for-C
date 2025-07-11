#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Sample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetFont.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section 
	intrusive_ptr<Section> s = doc->GetSections()->GetItemInSectionCollection(0);

	//Get the second paragraph
	intrusive_ptr<Paragraph> p = s->GetParagraphs()->GetItemInParagraphCollection(1);

	//Create a characterFormat object
	intrusive_ptr<CharacterFormat> format = new CharacterFormat(doc);
	//Set font
	format->SetFontName(L"Arial");
	format->SetFontSize(16);

	//Loop through the childObjects of paragraph 
	int pChildObjectsCount = p->GetChildObjects()->GetCount();
	for (int i = 0; i < pChildObjectsCount; i++)
	{
		intrusive_ptr<DocumentObject> childObj = p->GetChildObjects()->GetItem(i);
		if (Object::CheckType<TextRange>(childObj))
		{
			//Apply character format
			intrusive_ptr<TextRange> tr = boost::dynamic_pointer_cast<TextRange>(childObj);
			tr->ApplyCharacterFormat(format);
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}

