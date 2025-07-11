#include "pch.h"
#include <fstream>
#include <locale>
#include <codecvt>


using namespace Spire::Doc;

std::wstring GetBorderStyle(BorderStyle value)
{
	switch (value)
	{
	case Spire::Doc::BorderStyle::None:
		return L"None";
		break;
	case Spire::Doc::BorderStyle::Single:
		return L"Single";
		break;
	case Spire::Doc::BorderStyle::Thick:
		return L"Thick";
		break;
	case Spire::Doc::BorderStyle::Double:
		return L"Double";
		break;
	case Spire::Doc::BorderStyle::Hairline:
		return L"Hairline";
		break;
	case Spire::Doc::BorderStyle::Dot:
		return L"Dot";
		break;
	case Spire::Doc::BorderStyle::DashLargeGap:
		return L"DashLargeGap";
		break;
	case Spire::Doc::BorderStyle::DotDash:
		return L"DotDash";
		break;
	case Spire::Doc::BorderStyle::DotDotDash:
		return L"DotDotDash";
		break;
	case Spire::Doc::BorderStyle::Triple:
		return L"Triple";
		break;
	case Spire::Doc::BorderStyle::ThinThickSmallGap:
		return L"ThinThickSmallGap";
		break;
	case Spire::Doc::BorderStyle::ThickThinSmallGap:
		return L"ThickThinSmallGap";
		break;

	case Spire::Doc::BorderStyle::ThinThickThinSmallGap:
		return L"ThinThickThinSmallGap";
		break;
	case Spire::Doc::BorderStyle::ThinThickMediumGap:
		return L"ThinThickMediumGap";
		break;
	case Spire::Doc::BorderStyle::ThickThinMediumGap:
		return L"ThickThinMediumGap";
		break;
	case Spire::Doc::BorderStyle::ThinThickThinMediumGap:
		return L"ThinThickThinMediumGap";
		break;

	case Spire::Doc::BorderStyle::ThinThickLargeGap:
		return L"ThinThickLargeGap";
		break;
	case Spire::Doc::BorderStyle::ThickThinLargeGap:
		return L"ThickThinLargeGap";
		break;
	case Spire::Doc::BorderStyle::ThinThickThinLargeGap:
		return L"ThinThickThinLargeGap";
		break;
	case Spire::Doc::BorderStyle::Wave:
		return L"Wave";
		break;
	case Spire::Doc::BorderStyle::DoubleWave:
		return L"DoubleWave";
		break;
	case Spire::Doc::BorderStyle::DashSmallGap:
		return L"DashSmallGap";
		break;
	case Spire::Doc::BorderStyle::DashDotStroker:
		return L"DashDotStroker";
		break;
	case Spire::Doc::BorderStyle::Emboss3D:
		return L"Emboss3D";
		break;
	case Spire::Doc::BorderStyle::Engrave3D:
		return L"Engrave3D";
		break;
	case Spire::Doc::BorderStyle::Outset:
		return L"Outset";
		break;
	case Spire::Doc::BorderStyle::Inset:
		return L"Inset";
		break;
	case Spire::Doc::BorderStyle::TwistedLines1:
		return L"TwistedLines1";
		break;
	case Spire::Doc::BorderStyle::Cleared:
		return L"Cleared";
		break;
	}
	return L"";
}

int main()
{
	std::wstring outputFile = OUTPUTPATH"/GetDiagonalBorder.txt";
	std::wstring inputFile = DATAPATH"/GetDiagonalBorderOfCell.docx";
	
	//Load Word from disk
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	//Get the first table in the section
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

	std::wstring stringBuilder;

	//Get the setting of the diagonal border of table cell
	BorderStyle bs_UP = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetDiagonalUp()->GetBorderType();
	stringBuilder.append(L"DiagonalUp border type of table cell(0,0) is " + GetBorderStyle(bs_UP)).append(L"\n");

	intrusive_ptr<Color> color_UP = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetDiagonalUp()->GetColor();
	stringBuilder.append(L"DiagonalUp border color of table cell(0,0) is " + (std::wstring)color_UP->ToString()).append(L"\n");

	float width_UP = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetDiagonalUp()->GetLineWidth();
	stringBuilder.append(L"Line width of DiagonalUp border of table cell(0,0) is " + std::to_wstring(width_UP)).append(L"\n");

	BorderStyle bs_Down = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetDiagonalDown()->GetBorderType();
	stringBuilder.append(L"DiagonalDown border type of table cell(0,0) is " + GetBorderStyle(bs_Down)).append(L"\n");

	intrusive_ptr<Color> color_Down = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetDiagonalDown()->GetColor();
	stringBuilder.append(L"DiagonalDown border color of table cell(0,0) is " + (std::wstring)color_Down->ToString()).append(L"\n");

	float width_Down = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetDiagonalDown()->GetLineWidth();
	stringBuilder.append(L"DiagonalDown border color of table cell(0,0) is " + std::to_wstring(width_Down)).append(L"\n");

	//Save to txt file
	std::wofstream write(outputFile);
	auto LocUtf8 = locale(locale(""), new std::codecvt_utf8<wchar_t>);
	write.imbue(LocUtf8);
	write << stringBuilder;
	write.close();

	doc->Close();
}