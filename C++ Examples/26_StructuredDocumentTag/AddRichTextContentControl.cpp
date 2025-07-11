#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddRichTextContentControl.docx";

	//Create a document
	intrusive_ptr<Document> document = new Document();

	//Add a new section.
	intrusive_ptr<Section> section = document->AddSection();

	//Add a paragraph
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

	//Append textRange for the paragraph
	intrusive_ptr<TextRange> txtRange = paragraph->AppendText(L"The following example shows how to add RichText content control in a Word document. \n");

	//Append textRange 
	txtRange = paragraph->AppendText(L"Add RichText Content Control:  ");

	//Set the font format
	txtRange->GetCharacterFormat()->SetItalic(true);

	//Create StructureDocumentTagInline for document
	intrusive_ptr<StructureDocumentTagInline> sdt = new StructureDocumentTagInline(document);

	//Add sdt in paragraph
	paragraph->GetChildObjects()->Add(sdt);

	//Specify the type
	sdt->GetSDTProperties()->SetSDTType(SdtType::RichText);

	//Set displaying text
	intrusive_ptr<SdtText> text = new SdtText(true);
	text->SetIsMultiline(true);
	sdt->GetSDTProperties()->SetControlProperties(text);

	//Crate a TextRange
	intrusive_ptr<TextRange> rt = new TextRange(document);
	rt->SetText(L"Welcome to use ");
	rt->GetCharacterFormat()->SetTextColor(Color::GetGreen());
	sdt->GetSDTContent()->GetChildObjects()->Add(rt);

	rt = new TextRange(document);
	rt->SetText(L"Spire.Doc");
	rt->GetCharacterFormat()->SetTextColor(Color::GetOrangeRed());
	sdt->GetSDTContent()->GetChildObjects()->Add(rt);

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
