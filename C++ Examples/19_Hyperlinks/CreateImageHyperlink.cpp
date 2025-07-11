#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/CreateImageHyperlink.docx";
	std::wstring inputFile = DATAPATH"/BlankTemplate.docx";
	std::wstring inputFile_1 = DATAPATH"/Spire.Doc.png";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);
	//Add a paragraph
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
	//Load an image to a DocPicture object
#if defined(SKIASHARP)
	intrusive_ptr<DocPicture> picture = new DocPicture(doc);
	//Add an image hyperlink to the paragraph
	picture->LoadImageSpire(inputFile.c_str()_1);
#else
	intrusive_ptr<DocPicture> picture = new DocPicture(doc);
	//Add an image hyperlink to the paragraph
	picture->LoadImageSpire(inputFile_1.c_str());
#endif
	paragraph->AppendHyperlink(L"https://www.e-iceblue.com/Introduce/word-for-net-introduce.html", picture, HyperlinkType::WebLink);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}