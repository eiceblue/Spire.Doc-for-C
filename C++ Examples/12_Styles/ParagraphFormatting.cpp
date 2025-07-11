#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ParagraphFormatting.docx";

	//Initialize a document
	intrusive_ptr<Document> document = new Document();
	intrusive_ptr<Section> sec = document->AddSection();
	intrusive_ptr<Paragraph> para = sec->AddParagraph();
	para->AppendText(L"Paragraph Formatting");
	para->ApplyStyle(BuiltinStyle::Title);

	para = sec->AddParagraph();
	para->AppendText(L"This paragraph is surrounded with borders.");
	para->GetFormat()->GetBorders()->SetBorderType(BorderStyle::Single);
	para->GetFormat()->GetBorders()->SetColor(Color::GetRed());

	para = sec->AddParagraph();
	para->AppendText(L"The alignment of this paragraph is Left.");
	para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);

	para = sec->AddParagraph();
	para->AppendText(L"The alignment of this paragraph is Center.");
	para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

	para = sec->AddParagraph();
	para->AppendText(L"The alignment of this paragraph is Right.");
	para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

	para = sec->AddParagraph();
	para->AppendText(L"The alignment of this paragraph is justified.");
	para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Justify);

	para = sec->AddParagraph();
	para->AppendText(L"The alignment of this paragraph is distributed.");
	para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Distribute);

	para = sec->AddParagraph();
	para->AppendText(L"This paragraph has the gray shadow.");
	para->GetFormat()->SetBackColor(Color::GetGray());

	para = sec->AddParagraph();
	para->AppendText(L"This paragraph has the following indentations: Left indentation is 10pt, right indentation is 10pt, first line indentation is 15pt.");
	para->GetFormat()->SetLeftIndent(10);
	para->GetFormat()->SetRightIndent(10);
	para->GetFormat()->SetFirstLineIndent(15);

	para = sec->AddParagraph();
	para->AppendText(L"The hanging indentation of this paragraph is 15pt.");
	//Negative value represents hanging indentation
	para->GetFormat()->SetFirstLineIndent(-15);

	para = sec->AddParagraph();
	para->AppendText(L"This paragraph has the following spacing: spacing before is 10pt, spacing after is 20pt, line spacing is at least 10pt.");
	para->GetFormat()->SetAfterSpacing(20);
	para->GetFormat()->SetBeforeSpacing(10);
	para->GetFormat()->SetLineSpacingRule(LineSpacingRule::AtLeast);
	para->GetFormat()->SetLineSpacing(10);

	//Save as docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}