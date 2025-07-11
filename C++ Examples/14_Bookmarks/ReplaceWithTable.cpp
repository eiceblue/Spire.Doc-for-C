#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/ReplaceWithTable.docx";
	std::wstring inputFile = DATAPATH"/Bookmark.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Create a table
	intrusive_ptr<Table> table = new Table(doc, true);

	//Create data
	const int rowsCount = 4;
	const int colsCount = 5;
	std::wstring data[rowsCount][colsCount] = {
		{L"Name", L"Capital", L"Continent", L"Area", L"Population"},
		{L"Argentina", L"Buenos Aires", L"South America", L"2777815", L"32300003"},
		{L"Bolivia", L"La Paz", L"South America", L"1098575", L"7300000"},
		{L"Brazil", L"Brasilia", L"South America", L"8511196", L"150400000"}
	};
	/*data[0] = ["Name", "Capital", "Continent", "Area", "Population"];
	data[1] = ["Argentina", "Buenos Aires", "South America", "2777815", "32300003"];
	data[2] = ["Bolivia", "La Paz", "South America", "1098575", "7300000"];
	data[3] = ["Brazil", "Brasilia", "South America", "8511196", "150400000"];*/
	table->ResetCells(rowsCount, colsCount);

	//Fill the table with the data
	for (int i = 0; i < rowsCount; i++)
	{
		for (int j = 0; j < colsCount; j++)
		{
			table->GetRows()->GetItemInRowCollection(i)->GetCells()->GetItemInCellCollection(j)->AddParagraph()->AppendText(data[i][j].c_str());
		}
	}

	//Get the specific bookmark by its name
	intrusive_ptr<BookmarksNavigator> navigator = new BookmarksNavigator(doc);
	navigator->MoveToBookmark(L"Test");

	//Create a TextGetBody()Part instance and add the table to it
	intrusive_ptr<TextBodyPart> part = new TextBodyPart(doc);
	part->GetBodyItems()->Add(table);

	//Replace the current bookmark content with the TextGetBody()Part object
	navigator->ReplaceBookmarkContent(part);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();

}