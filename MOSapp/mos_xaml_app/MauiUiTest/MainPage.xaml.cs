using MOSExcelMogiApp.Maui.Pages;

namespace MOSExcelMogiApp.Maui;

public partial class MainPage : ContentPage
{
	public MainPage()
	{
		InitializeComponent();
	}

	private async void UiTestButton_Clicked(object? sender, EventArgs e)
	{
		// UIテストページに遷移
		await Navigation.PushAsync(new UiTestAppBarPage());
	}
}
