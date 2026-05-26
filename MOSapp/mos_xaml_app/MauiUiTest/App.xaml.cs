using System;
using System.Threading.Tasks;
using Microsoft.Maui.Devices;
using Microsoft.Maui.Controls;

namespace MOSExcelMogiApp.Maui;

public partial class App : Application
{
	public App()
	{
		InitializeComponent();
	}

	protected override Window CreateWindow(IActivationState? activationState)
	{
		var window = new Window(new AppShell());
		
#if WINDOWS
		// Windowsプラットフォームでウィンドウのサイズと位置を設定
		window.Created += (sender, args) =>
		{
			SetWindowSizeAndPosition(window);
		};
		
		// Handlerが作成された後にも設定
		window.HandlerChanged += (sender, args) =>
		{
			if (window.Handler != null)
			{
				Task.Delay(100).ContinueWith(_ =>
				{
					MainThread.BeginInvokeOnMainThread(() =>
					{
						SetWindowSizeAndPosition(window);
					});
				});
			}
		};
#endif
		
		return window;
	}
	
#if WINDOWS
	private void SetWindowSizeAndPosition(Window window)
	{
		try
		{
			var displayInfo = DeviceDisplay.MainDisplayInfo;
			var screenHeight = displayInfo.Height / displayInfo.Density;
			var screenWidth = displayInfo.Width / displayInfo.Density;
			
			// 画面の1/3の高さに設定（下1/3まで表示）
			var targetHeight = screenHeight / 3.0;
			var targetY = screenHeight * 2.0 / 3.0;
			
			if (window.Handler?.PlatformView is Microsoft.UI.Xaml.Window platformWindow)
			{
				var size = new Windows.Graphics.SizeInt32((int)screenWidth, (int)targetHeight);
				platformWindow.AppWindow.Resize(size);
				
				var displayArea = Microsoft.UI.Windowing.DisplayArea.GetFromWindowId(
					platformWindow.AppWindow.Id, 
					Microsoft.UI.Windowing.DisplayAreaFallback.Nearest);
				if (displayArea != null)
				{
					var x = displayArea.WorkArea.X;
					var y = (int)targetY;
					platformWindow.AppWindow.Move(new Windows.Graphics.PointInt32(x, y));
					
					System.Diagnostics.Debug.WriteLine($"App: Window resized to {screenWidth}x{targetHeight} at position ({x}, {y})");
				}
			}
		}
		catch (Exception ex)
		{
			System.Diagnostics.Debug.WriteLine($"SetWindowSizeAndPosition error: {ex.Message}");
		}
	}
#endif
}