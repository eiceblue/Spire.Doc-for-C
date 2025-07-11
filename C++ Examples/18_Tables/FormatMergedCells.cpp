#include "pch.h"

using namespace Spire::Doc;

intrusive_ptr<Table> AddTable(intrusive_ptr<Section> section)
{
	intrusive_ptr<Table> table = section->AddTable(true);
	table->ResetCells(4, 3);
	std::vector<std::vector<std::wstring>> data =
	{
		{L"Product", L"", L"Price"},
		{L"Spire.Doc", L"Pro Edition", L"$799"},
		{L"", L"Standard Edition", L"$599"},
		{L"", L"Free Edition", L"$0"}
	};
	for (size_t r = 0; r < data.size(); r++)
	{
		intrusive_ptr<TableRow> dataRow = table->GetRows()->GetItemInRowCollection(r);
		dataRow->SetHeight(20);
		dataRow->SetHeightType(TableRowHeightType::Exactly);

		for (size_t c = 0; c < data[r].size(); c++)
		{
			if (!data[r][c].empty()) {
				intrusive_ptr<TextRange> range = dataRow->GetCells()->GetItemInCellCollection(c)->AddParagraph()->AppendText(data[r][c].c_str());
				range->GetCharacterFormat()->SetFontName(L"Arial");

				dataRow->GetCells()->GetItemInCellCollection(c)->GetCellFormat()->GetShading()->SetBackgroundPatternColor(Color::Empty());
			}
		}
	}
	return table;
}

int main()
{
	std::wstring outputFile = OUTPUTPATH"/FormatMergedCells.docx";

	//Create word document
	intrusive_ptr<Document> document = new Document();

	//Add a new section
	intrusive_ptr<Section> section = document->AddSection();

	//Add a new table
	intrusive_ptr<Table> table = AddTable(section);

	//Create a new style
	intrusive_ptr<ParagraphStyle> style = new ParagraphStyle(document);
	style->SetName(L"Style");
	style->GetCharacterFormat()->SetTextColor(Color::GetDeepSkyBlue());
	style->GetCharacterFormat()->SetItalic(true);
	style->GetCharacterFormat()->SetBold(true);
	style->GetCharacterFormat()->SetFontSize(13);
	document->GetStyles()->Add(style);

	//Merge cell horizontally
	table->ApplyHorizontalMerge(0, 0, 1);
	//Apply style
	table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0)->ApplyStyle(style->GetName());
	//Set vertical and horizontal alignment
	table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
	table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0)->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

	//Merge cell vertically
	table->ApplyVerticalMerge(0, 1, 3);
	//Apply style
	table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0)->ApplyStyle(style->GetName());
	//Set vertical and horizontal alignment
	table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
	table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0)->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);
	//Set column width
	table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(0)->SetCellWidth(20, CellWidthType::Percentage);
	//Save and launch document
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();

}