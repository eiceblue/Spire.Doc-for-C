#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/SetTableStyleAndBorder.docx";
	std::wstring inputFile = DATAPATH"/TableSample.docx";

	//Create a document and load file
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

	//Get the first table
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

	//Apply the table style
	table->ApplyStyle(DefaultTableStyle::ColorfulList);

	//Set right border of table
	table->GetFormat()->GetBorders()->GetRight()->SetBorderType(BorderStyle::Hairline);
	table->GetFormat()->GetBorders()->GetRight()->SetLineWidth(1.0F);
	table->GetFormat()->GetBorders()->GetRight()->SetColor(Color::GetRed());

	//Set top border of table
	table->GetFormat()->GetBorders()->GetTop()->SetBorderType(BorderStyle::Hairline);
	table->GetFormat()->GetBorders()->GetTop()->SetLineWidth(1.0F);
	table->GetFormat()->GetBorders()->GetTop()->SetColor(Color::GetGreen());

	//Set left border of table
	table->GetFormat()->GetBorders()->GetLeft()->SetBorderType(BorderStyle::Hairline);
	table->GetFormat()->GetBorders()->GetLeft()->SetLineWidth(1.0F);
	table->GetFormat()->GetBorders()->GetLeft()->SetColor(Color::GetYellow());

	//Set bottom border is none
	table->GetFormat()->GetBorders()->GetBottom()->SetBorderType(BorderStyle::DotDash);

	//Set vertical and horizontal border 
	table->GetFormat()->GetBorders()->GetVertical()->SetBorderType(BorderStyle::Dot);
	table->GetFormat()->GetBorders()->GetHorizontal()->SetBorderType(BorderStyle::None);
	table->GetFormat()->GetBorders()->GetVertical()->SetColor(Color::GetOrange());

	//Save the file and launch it
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();

}
