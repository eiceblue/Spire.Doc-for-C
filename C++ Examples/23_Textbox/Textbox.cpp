#include "pch.h"
using namespace Spire::Doc;

void InsertTextbox(intrusive_ptr<Section> section)
{
	intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetCount() > 0 ? section->GetParagraphs()->GetItemInParagraphCollection(0) : section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();

	//Insert and format the first textbox.
	intrusive_ptr<TextBox> textBox1 = paragraph->AppendTextBox(240, 35);
	textBox1->GetFormat()->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	textBox1->GetFormat()->SetLineColor(Color::GetGray());
	textBox1->GetFormat()->SetLineStyle(TextBoxLineStyle::Simple);
	textBox1->GetFormat()->SetFillColor(Color::GetDarkSeaGreen());
	intrusive_ptr<Paragraph> para = textBox1->GetBody()->AddParagraph();
	intrusive_ptr<TextRange> txtrg = para->AppendText(L"Textbox 1 in the document");
	txtrg->GetCharacterFormat()->SetFontName(L"Lucida Sans Unicode");
	txtrg->GetCharacterFormat()->SetFontSize(14);
	txtrg->GetCharacterFormat()->SetTextColor(Color::GetWhite());
	para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

	//Insert and format the second textbox.
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	intrusive_ptr<TextBox> textBox2 = paragraph->AppendTextBox(240, 35);
	textBox2->GetFormat()->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	textBox2->GetFormat()->SetLineColor(Color::GetTomato());
	textBox2->GetFormat()->SetLineStyle(TextBoxLineStyle::ThinThick);
	textBox2->GetFormat()->SetFillColor(Color::GetBlue());
	textBox2->GetFormat()->SetLineDashing(LineDashing::Dot);
	para = textBox2->GetBody()->AddParagraph();
	txtrg = para->AppendText(L"Textbox 2 in the document");
	txtrg->GetCharacterFormat()->SetFontName(L"Lucida Sans Unicode");
	txtrg->GetCharacterFormat()->SetFontSize(14);
	txtrg->GetCharacterFormat()->SetTextColor(Color::GetPink());
	para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

	//Insert and format the third textbox.
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	intrusive_ptr<TextBox> textBox3 = paragraph->AppendTextBox(240, 35);
	textBox3->GetFormat()->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	textBox3->GetFormat()->SetLineColor(Color::GetViolet());
	textBox3->GetFormat()->SetLineStyle(TextBoxLineStyle::Triple);
	textBox3->GetFormat()->SetFillColor(Color::GetPink());
	textBox3->GetFormat()->SetLineDashing(LineDashing::DashDotDot);
	para = textBox3->GetBody()->AddParagraph();
	txtrg = para->AppendText(L"Textbox 3 in the document");
	txtrg->GetCharacterFormat()->SetFontName(L"Lucida Sans Unicode");
	txtrg->GetCharacterFormat()->SetFontSize(14);
	txtrg->GetCharacterFormat()->SetTextColor(Color::GetTomato());
	para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
}
int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Textbox.docx";

	//Create a Word document and and a section.
	intrusive_ptr<Document> document = new Document();
	intrusive_ptr<Section> section = document->AddSection();

	InsertTextbox(section);

	//Save docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}


