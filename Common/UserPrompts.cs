using DocumentFormat.OpenXml.Office2019.Excel.RichData;
using Spectre.Console;
using System;

public static class UserPrompts
{
    // Wenn true, wird die Nachfrage bis zum Neustart der App unterdrückt
    private static bool _skipConfirmationSeiten = false;
    private static bool _skipConfirmationDateien = false;

    // Anzahl an künftigen Durchläufen, die ohne Nachfrage übersprungen werden sollen
    private static int _skipConfirmationSeitenCount = 0;
    private static int _skipConfirmationDateienCount = 0;

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
    /// ENTER = weiter, x = Exception (Abbruch), w = nicht mehr fragen in dieser Session,
    /// Ziffer 1-9 = 10-90 weitere Durchläufe ohne Nachfrage.
    /// </summary>
    public static void ConfirmOrThrowSeiten(string question)
    {
        // Wenn dauerhaft unterdrückt, sofort zurück
        if (_skipConfirmationSeiten) return;

        // Wenn temporär überspringen (Zähler > 0), dann Zähler verringern und zurück
        if (_skipConfirmationSeitenCount > 0)
        {
            _skipConfirmationSeitenCount--;
            return;
        }

        AnsiConsole.MarkupLine($"[{Global.GetColor(Global.ColorHinweise)}]{question}[/]");
        AnsiConsole.MarkupLine($"Drücken Sie [green]ENTER[/] (weiter), [red]x[/] (abbrechen), [yellow]w[/] (weiter, nicht mehr fragen in dieser Session) oder eine Ziffer [aqua]1-9[/] (10-90 Durchläufe überspringen).");

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

            // Ziffern 1-9: setze temporären Skip-Zähler (10-90 Durchläufe)
            if (c >= '1' && c <= '9')
            {
                int multiplier = c - '0';
                _skipConfirmationSeitenCount = multiplier * 10;
                AnsiConsole.MarkupLine($"[yellow]Überspringe die nächsten {_skipConfirmationSeitenCount} Durchläufe ohne Nachfrage.[/]");
                return;
            }
            // ungültige Taste → erneut fragen
        }
    }

    /// <summary>
    /// Fragt den Benutzer mit einer individuellen Frage.
    /// ENTER = weiter, x = Exception (Abbruch), w = nicht mehr fragen in dieser Session,
    /// Ziffer 1-9 = 10-90 weitere Durchläufe ohne Nachfrage.
    /// </summary>
    public static void ConfirmOrThrowDateien(string question)
    {
        // Wenn dauerhaft unterdrückt, sofort zurück
        if (_skipConfirmationDateien) return;

        // Wenn temporär überspringen (Zähler > 0), dann Zähler verringern und zurück
        if (_skipConfirmationDateienCount > 0)
        {
            _skipConfirmationDateienCount--;
            return;
        }

        AnsiConsole.MarkupLine($"[{Global.GetColor(Global.ColorHinweise)}]{question}[/]");
        AnsiConsole.MarkupLine($"Drücken Sie [green]ENTER[/] (weiter), [red]x[/] (abbrechen), [yellow]w[/] (weiter, nicht mehr fragen in dieser Session) oder eine Ziffer [aqua]1-9[/] (10-90 Durchläufe überspringen).");

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

            // Ziffern 1-9: setze temporären Skip-Zähler (10-90 Durchläufe)
            if (c >= '1' && c <= '9')
            {
                int multiplier = c - '0';
                _skipConfirmationDateienCount = multiplier * 10;
                AnsiConsole.MarkupLine($"[yellow]Überspringe die nächsten {_skipConfirmationDateienCount} Durchläufe ohne Nachfrage.[/]");
                return;
            }
            // ungültige Taste → erneut fragen
        }
    }
}
