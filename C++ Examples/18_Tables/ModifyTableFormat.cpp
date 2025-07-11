#include "pch.h"

using namespace Spire::Doc;

void MoidyTableFormat(intrusive_ptr<Table> table)
{
	//Set table width
	table->SetPreferredWidth(new PreferredWidth(WidthType::Twip, static_cast<short>(6000)));

	//Apply style for table
	table->ApplyStyle(DefaultTableStyle::ColorfulGridAccent3);

	//Set table padding
	table->GetFormat()->GetPaddings()->SetAll(5);

	//Set table title and description
	table->SetTitle(L"Spire.Doc for .NET");
	table->SetTableDescription(L"Spire.Doc for .NET is a professional Word .NET library");
}

void ModifyRowFormat(intrusive_ptr<Table> table)
{
	//Set cell spacing
	table->GetFormat()->SetCellSpacing(2);

	//Set row height
	table->GetRows()->GetItemInRowCollection(1)->SetHeightType(TableRowHeightType::Exactly);
	table->GetRows()->GetItemInRowCollection(1)->SetHeight(20.0f);

	//Set background color
	for (int i = 0; i < table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetCount(); i++)
	{
		table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(i)->GetCellFormat()->GetShading()->SetBackgroundPatternColor(Color::GetDarkSeaGreen());
	}
}

void ModifyCellFormat(intrusive_ptr<Table> table)
{
	//Set alignment
	table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
	table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0)->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

	//Set background color
	table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetShading()->SetBackgroundPatternColor(Color::GetDarkSeaGreen());

	//Set cell border
	table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->SetBorderType(BorderStyle::Single);
	table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->SetLineWidth(1.0f);
	table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetLeft()->SetColor(Color::GetRed());
	table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetRight()->SetColor(Color::GetRed());
	table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetTop()->SetColor(Color::GetRed());
	table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetBottom()->SetColor(Color::GetRed());

	//Set text direction
	table->GetRows()->GetItemInRowCollection(3)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->SetTextDirection(TextDirection::RightToLeft);
}

int main()
{
	std::wstring outputFile = OUTPUTPATH"/ModifyTableFormat.docx";
	std::wstring inputFile = DATAPATH"/ModifyTableFormat.docx";

	//Load Word document from disk
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

	//Get tables 
	intrusive_ptr<Table> tb1 = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));
	intrusive_ptr<Table> tb2 = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(1));
	intrusive_ptr<Table> tb3 = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(2));

	MoidyTableFormat(tb1);
	ModifyRowFormat(tb2);
	ModifyCellFormat(tb3);

	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}