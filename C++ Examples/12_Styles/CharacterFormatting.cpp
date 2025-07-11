#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CharacterFormatting.docx";

	//Initialize a document
	intrusive_ptr<Document> document = new Document();
	intrusive_ptr<Section> sec = document->AddSection();
	intrusive_ptr<Paragraph> titleParagraph = sec->AddParagraph();
	titleParagraph->AppendText(L"Font Styles and Effects ");
	titleParagraph->ApplyStyle(BuiltinStyle::Title);

	intrusive_ptr<Paragraph> paragraph = sec->AddParagraph();
	intrusive_ptr<TextRange> tr = paragraph->AppendText(L"Strikethough Text");
	tr->GetCharacterFormat()->SetIsStrikeout(true);

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"Shadow Text");
	tr->GetCharacterFormat()->SetIsShadow(true);

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"Small caps Text");
	tr->GetCharacterFormat()->SetIsSmallCaps(true);

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"Double Strikethough Text");
	tr->GetCharacterFormat()->SetDoubleStrike(true);

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"Outline Text");
	tr->GetCharacterFormat()->SetIsOutLine(true);

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"AllCaps Text");
	tr->GetCharacterFormat()->SetAllCaps(true);

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"Text");
	tr = paragraph->AppendText(L"SubScript");
	tr->GetCharacterFormat()->SetSubSuperScript(SubSuperScript::SubScript);

	tr = paragraph->AppendText(L"And");
	tr = paragraph->AppendText(L"SuperScript");
	tr->GetCharacterFormat()->SetSubSuperScript(SubSuperScript::SuperScript);

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"Emboss Text");
	tr->GetCharacterFormat()->SetEmboss(true);
	tr->GetCharacterFormat()->SetTextColor(Color::GetWhite());

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"Hidden:");
	tr = paragraph->AppendText(L"Hidden Text");
	tr->GetCharacterFormat()->SetHidden(true);

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"Engrave Text");
	tr->GetCharacterFormat()->SetEngrave(true);
	tr->GetCharacterFormat()->SetTextColor(Color::GetWhite());

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"WesternFonts中文字体");
	tr->GetCharacterFormat()->SetFontNameAscii(L"Calibri");
	tr->GetCharacterFormat()->SetFontNameNonFarEast(L"Calibri");
	tr->GetCharacterFormat()->SetFontNameFarEast(L"Simsun");

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"Font Size");
	tr->GetCharacterFormat()->SetFontSize(20);

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"Font Color");
	tr->GetCharacterFormat()->SetTextColor(Color::GetRed());

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"Bold Italic Text");
	tr->GetCharacterFormat()->SetBold(true);
	tr->GetCharacterFormat()->SetItalic(true);

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"Underline Style");
	tr->GetCharacterFormat()->SetUnderlineStyle(UnderlineStyle::Single);

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"Highlight Text");
	tr->GetCharacterFormat()->SetHighlightColor(Color::GetYellow());

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"Text has shading");
	tr->GetCharacterFormat()->SetTextBackgroundColor(Color::GetGreen());

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"Border Around Text");
	tr->GetCharacterFormat()->GetBorder()->SetBorderType(BorderStyle::Single);

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"Text Scale");
	tr->GetCharacterFormat()->SetTextScale(150);

	paragraph->AppendBreak(BreakType::LineBreak);
	tr = paragraph->AppendText(L"Character Spacing is 2 point");
	tr->GetCharacterFormat()->SetCharacterSpacing(2);

	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();

}
