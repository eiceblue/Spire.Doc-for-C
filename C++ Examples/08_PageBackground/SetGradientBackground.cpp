#include "pch.h"
using namespace Spire::Doc;


int main() {

	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_2.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetGradientBackground.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Set the background type as Gradient.
	document->GetBackground()->SetType(BackgroundType::Gradient);
	intrusive_ptr<BackgroundGradient> Test = document->GetBackground()->GetGradient();

	//Set the first color and second color for Gradient.
	Test->SetColor1(Color::GetWhite());
	Test->SetColor2(Color::GetLightBlue());

	//Set the Shading style and Variant for the gradient.
	Test->SetShadingVariant(GradientShadingVariant::ShadingDown);
	Test->SetShadingStyle(GradientShadingStyle::Horizontal);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();

}

