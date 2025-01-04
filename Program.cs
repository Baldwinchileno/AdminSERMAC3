﻿using Microsoft.Extensions.DependencyInjection;
using AdminSERMAC.Core.Configuration;
using AdminSERMAC.Services;
using AdminSERMAC.Core.Theme;
using AdminSERMAC.Core.Interfaces;

namespace AdminSERMAC;

static class Program
{
    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();

        // Inicializar el tema
        ThemeManager.LoadThemePreference();

        // Configurar servicios
        var services = new ServiceCollection();
        var connectionString = "Data Source=AdminSERMAC.db;Version=3;";

        services.AddInfrastructure(connectionString);

        var serviceProvider = services.BuildServiceProvider();

        var clienteService = serviceProvider.GetRequiredService<IClienteService>();

        Application.Run(new MainForm(clienteService, connectionString));
    }
}