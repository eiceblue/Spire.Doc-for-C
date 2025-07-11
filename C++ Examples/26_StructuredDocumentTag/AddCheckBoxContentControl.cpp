#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddCheckBoxContentControl.docx";

	//Create a document
	intrusive_ptr<Document> document = new Document();

	intrusive_ptr<Section> section = document->AddSection();

	//Add a paragraph
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

	//Append textRange for the paragraph
	intrusive_ptr<TextRange> txtRange = paragraph->AppendText(L"The following example shows how to add CheckBox content control in a Word document. \n");

	//Append textRange 
	txtRange = paragraph->AppendText(L"Add CheckBox Content Control:  ");

	//Set the font format
	txtRange->GetCharacterFormat()->SetItalic(true);

	//Create StructureDocumentTagInline for document
	intrusive_ptr<StructureDocumentTagInline> sdt = new StructureDocumentTagInline(document);

	//Add sdt in paragraph
	paragraph->GetChildObjects()->Add(sdt);

	//Specify the type
	sdt->GetSDTProperties()->SetSDTType(SdtType::CheckBox);

	//Set properties for control
	intrusive_ptr<SdtCheckBox> scb = new SdtCheckBox();
	sdt->GetSDTProperties()->SetControlProperties(scb);

	//Add textRange format
	intrusive_ptr<TextRange> tr = new TextRange(document);
	tr->GetCharacterFormat()->SetFontName(L"MS Gothic");
	tr->GetCharacterFormat()->SetFontSize(12);

	//Add textRange to StructureDocumentTagInline
	sdt->GetChildObjects()->Add(tr);

	//Set checkBox as checked
	scb->SetChecked(true);

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
