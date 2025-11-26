using DocumentFormat.OpenXml.Office2019.Excel.RichData;
using Spectre.Console;
using System;

public static class UserPrompts
{
    // Wenn true, wird die Nachfrage bis zum Neustart der App unterdrückt
    private static bool _skipConfirmationSeiten = false;
    private static bool _skipConfirmationDateien = false;

    public static bool SkipConfirmationSeiten
    {
        get => _skipConfirmationSeiten;
        set => _skipConfirmationSeiten = value;
    }

    public static bool SkipConfirmationDateien
    {
        get => _skipConfirmationDateien;
        set => _skipConfirmationDateien = value;
    }

    /// <summary>
    /// Fragt den Benutzer mit einer individuellen Frage.
    /// j = weiter, x = Exception (Abbruch), w = nicht mehr fragen in dieser Session.
    /// </summary>
    public static void ConfirmOrThrowSeiten(string question)
    {
        if (_skipConfirmationSeiten) return;     

        AnsiConsole.MarkupLine($"[{Global.GetColor(Global.ColorHinweise)}]{question}[/]");
        AnsiConsole.MarkupLine($"Drücken Sie [green]ENTER[/] (weiter), [red]x[/] (abbrechen) oder [yellow]w[/] (weiter, nicht mehr fragen in dieser Session).");

        while (true)
        {
            var keyInfo = Console.ReadKey(true);

            // Fall back, falls keine gültige Taste gelesen wurde
            if (keyInfo.KeyChar == 0) continue;

            var c = char.ToLowerInvariant(keyInfo.KeyChar);

            if (c == 'x')
            {
                throw new OperationCanceledException("Sie haben abgebrochen.");
            }

            if (c == 'j' || keyInfo.Key == ConsoleKey.Enter)
            {
                return; // bestätigen und weiter
            }

            if (c == 'w')
            {
                _skipConfirmationSeiten = true;
                return; // bestätigen und künftige Fragen unterdrücken
            }
        }
    }

    /// <summary>
    /// Fragt den Benutzer mit einer individuellen Frage.
    /// j = weiter, x = Exception (Abbruch), w = nicht mehr fragen in dieser Session.
    /// </summary>
    public static void ConfirmOrThrowDateien(string question)
    {   
        if (_skipConfirmationDateien) return;

        AnsiConsole.MarkupLine($"[{Global.GetColor(Global.ColorHinweise)}]{question}[/]");
        AnsiConsole.MarkupLine($"Drücken Sie [green]ENTER[/] (weiter), [red]x[/] (abbrechen) oder [yellow]w[/] (weiter, nicht mehr fragen in dieser Session).");

        while (true)
        {
            var keyInfo = Console.ReadKey(true);

            // Fall back, falls keine gültige Taste gelesen wurde
            if (keyInfo.KeyChar == 0) continue;

            var c = char.ToLowerInvariant(keyInfo.KeyChar);

            if (c == 'x')
            {
                throw new OperationCanceledException("Sie haben abgebrochen.");
            }

            if (c == 'j' || keyInfo.Key == ConsoleKey.Enter)
            {
                return; // bestätigen und weiter
            }

            if (c == 'w')
            {
                _skipConfirmationDateien = true;
                return; // bestätigen und künftige Fragen unterdrücken
            }
        }
    }
}
