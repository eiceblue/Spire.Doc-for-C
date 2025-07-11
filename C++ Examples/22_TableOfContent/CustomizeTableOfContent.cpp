#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CustomizeTableOfContent.docx";

	//Create a document
	intrusive_ptr<Document> doc = new Document();
	//Add a section
	intrusive_ptr<Section> section = doc->AddSection();
	//Customize table of contents with switches
	intrusive_ptr<TableOfContent> toc = new TableOfContent(doc, L"{\\o \"1-3\" \\n 1-1}");
	intrusive_ptr<Paragraph> para = section->AddParagraph();
	para->GetItems()->Add(toc);
	para->AppendFieldMark(FieldMarkType::FieldSeparator);
	para->AppendText(L"TOC");
	para->AppendFieldMark(FieldMarkType::FieldEnd);
	doc->SetTOC(toc);

	intrusive_ptr<Paragraph> par = section->AddParagraph();
	intrusive_ptr<TextRange> tr = par->AppendText(L"Flowers");
	tr->GetCharacterFormat()->SetFontSize(30);
	par->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

	//Create paragraph and set the head level
	intrusive_ptr<Paragraph> para1 = section->AddParagraph();
	para1->AppendText(L"Ornithogalum");
	//Apply the Heading1 style
	para1->ApplyStyle(BuiltinStyle::Heading1);
	//Add paragraphs
	para1 = section->AddParagraph();
	intrusive_ptr<DocPicture> picture = para1->AppendPicture(DATAPATH"/Ornithogalum.jpg");

	picture->SetTextWrappingStyle(TextWrappingStyle::Square);
	para1->AppendText(L"Ornithogalum is a genus of perennial plants mostly native to southern Europe and southern Africa belonging to the family Asparagaceae. Some species are native to other areas such as the Caucasus. Growing from a bulb, species have linear basal leaves and a slender stalk, up to 30 cm tall, bearing clusters of typically white star-shaped flowers, often striped with green.");
	para1 = section->AddParagraph();

	intrusive_ptr<Paragraph> para2 = section->AddParagraph();
	para2->AppendText(L"Rosa");
	//Apply the Heading2 style
	para2->ApplyStyle(BuiltinStyle::Heading2);
	para2 = section->AddParagraph();
	intrusive_ptr<DocPicture> picture2 = para2->AppendPicture(DATAPATH"/Rosa.jpg");
	picture2->SetTextWrappingStyle(TextWrappingStyle::Square);
	para2->AppendText(L"A rose is a woody perennial flowering plant of the genus Rosa, in the family Rosaceae, or the flower it bears. There are over a hundred species and thousands of cultivars. They form a group of plants that can be erect shrubs, climbing or trailing with stems that are often armed with sharp prickles. Flowers vary in size and shape and are usually large and showy, in colours ranging from white through yellows and reds. Most species are native to Asia, with smaller numbers native to Europe, North America, and northwestern Africa. Species, cultivars and hybrids are all widely grown for their beauty and often are fragrant. Roses have acquired cultural significance in many societies. Rose plants range in size from compact, miniature roses, to climbers that can reach seven meters in height. Different species hybridize easily, and this has been used in the development of the wide range of garden roses.");
	section->AddParagraph();

	intrusive_ptr<Paragraph> para3 = section->AddParagraph();
	para3->AppendText(L"Hyacinth");
	//Apply the Heading3 style
	para3->ApplyStyle(BuiltinStyle::Heading3);
	para3 = section->AddParagraph();
	intrusive_ptr<DocPicture> picture3 = para3->AppendPicture(DATAPATH"/hyacinths.JPG");
	picture3->SetTextWrappingStyle(TextWrappingStyle::Tight);
	para3->AppendText(L"Hyacinthus is a small genus of bulbous, fragrant flowering plants in the family Asparagaceae, subfamily Scilloideae.These are commonly called hyacinths.The genus is native to the eastern Mediterranean (from the south of Turkey through to northern Israel).");
	para3 = section->AddParagraph();
	para3->AppendText(L"Several species of Brodiea, Scilla, and other plants that were formerly classified in the lily family and have flower clusters borne along the stalk also have common names with the word \"hyacinth\" in them. Hyacinths should also not be confused with the genus Muscari, which are commonly known as grape hyacinths.");
	
	//Update TOC
	doc->UpdateTableOfContents();
	//Save to file
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}