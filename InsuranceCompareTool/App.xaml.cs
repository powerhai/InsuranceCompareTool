﻿using InsuranceCompareTool.Views;
using Prism.Ioc;
using Prism.Modularity;
using System.Windows;
using InsuranceCompareTool.Services;
namespace InsuranceCompareTool
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App
    {
        protected override Window CreateShell()
        {
            return Container.Resolve<MainWindow>();
        }

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            var configService = this.Container.Resolve<ConfigService>(); 
            containerRegistry.RegisterInstance<ConfigService>(configService);
        }
    }
}
