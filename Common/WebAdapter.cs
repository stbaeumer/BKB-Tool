using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace Common
{
    public static class WebAdapter
    {
        // Erweitert die SchuelerZusatzdaten-Datei-Inhalte um schulinterne Mailadressen.
        public static string ErgaenzeSchuelerZusatzdatenMailadressen(string inputContent, string mailDomain)
        {
            if (string.IsNullOrWhiteSpace(inputContent))
                throw new ArgumentException("Input darf nicht leer sein.", nameof(inputContent));

            mailDomain ??= string.Empty;

            var lines = inputContent.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n').ToList();
            if (lines.Count == 0) return inputContent;

            int headerIndex = lines.FindIndex(l => !string.IsNullOrWhiteSpace(l));
            if (headerIndex < 0) return inputContent;

            var headerLine = lines[headerIndex];
            var headers = ParseLine(headerLine, '|', '\'').ToList();

            int idxNachname = headers.FindIndex(h => string.Equals(h, "Nachname", StringComparison.OrdinalIgnoreCase));
            int idxVorname = headers.FindIndex(h => string.Equals(h, "Vorname", StringComparison.OrdinalIgnoreCase));
            int idxGeb = headers.FindIndex(h => string.Equals(h, "Geburtsdatum", StringComparison.OrdinalIgnoreCase));

            var possibleMailCols = new[] { "MailSchulisch", "EMail", "EMINUSMail", "Email", "Mail", "schulische E-Mail" };
            int idxMail = headers.Select((h, i) => new { h, i })
                                 .Where(x => possibleMailCols.Any(p => string.Equals(p, x.h, StringComparison.OrdinalIgnoreCase)))
                                 .Select(x => x.i)
                                 .FirstOrDefault(-1);

            if (idxMail == -1)
            {
                headers.Add("schulische E-Mail");
                idxMail = headers.Count - 1;
            }

            var outputLines = new List<string>();
            outputLines.Add(JoinLine(headers, '|', '\''));

            for (int i = headerIndex + 1; i < lines.Count; i++)
            {
                var line = lines[i];
                if (string.IsNullOrWhiteSpace(line))
                {
                    outputLines.Add(line);
                    continue;
                }

                var fields = ParseLine(line, '|', '\'').ToList();
                if (fields.Count < headers.Count)
                {
                    while (fields.Count < headers.Count) fields.Add(string.Empty);
                }

                string nachname = idxNachname >= 0 && idxNachname < fields.Count ? fields[idxNachname] : "";
                string vorname = idxVorname >= 0 && idxVorname < fields.Count ? fields[idxVorname] : "";
                string geb = idxGeb >= 0 && idxGeb < fields.Count ? fields[idxGeb] : "";

                string existierendeMail = idxMail >= 0 && idxMail < fields.Count ? fields[idxMail] : "";

                if (string.IsNullOrWhiteSpace(existierendeMail))
                {
                    var gebPart = ParseGeburtsdatumZuYYMMDD(geb);
                    var n = BereinigenNachnameFuerMail(nachname);
                    var v = BereinigenVornameFuerMail(vorname);

                    if (!string.IsNullOrEmpty(n) && !string.IsNullOrEmpty(v) && !string.IsNullOrEmpty(gebPart) && !string.IsNullOrEmpty(mailDomain))
                    {
                        var neu = (n.Substring(0, 1) + v.Substring(0, 1) + gebPart + mailDomain).ToLowerInvariant();
                        fields[idxMail] = neu;
                    }
                }

                outputLines.Add(JoinLine(fields, '|', '\''));
            }

            return string.Join("\n", outputLines);
        }

        private static IEnumerable<string> ParseLine(string line, char delimiter, char quote)
        {
            var fields = new List<string>();
            if (line == null) return fields;

            var sb = new StringBuilder();
            bool inQuote = false;

            for (int i = 0; i < line.Length; i++)
            {
                var c = line[i];
                if (c == quote)
                {
                    inQuote = !inQuote;
                    continue;
                }

                if (c == delimiter && !inQuote)
                {
                    fields.Add(sb.ToString());
                    sb.Clear();
                    continue;
                }

                sb.Append(c);
            }

            fields.Add(sb.ToString());
            return fields.Select(f => f.Trim());
        }

        private static string JoinLine(IEnumerable<string> fields, char delimiter, char quote)
        {
            var outFields = fields.Select(f =>
            {
                if (f == null) return "";
                var needsQuote = f.Contains(delimiter) || f.Contains('\n') || f.Contains('\r') || f.Contains(quote);
                var escaped = f.Replace(quote.ToString(), new string(quote, 2));
                return needsQuote ? $"{quote}{escaped}{quote}" : escaped;
            });
            return string.Join(delimiter, outFields);
        }

        private static string ParseGeburtsdatumZuYYMMDD(string geb)
        {
            if (string.IsNullOrWhiteSpace(geb)) return "";
            geb = geb.Trim().Replace("'", "").Replace("\"", "");
            DateTime dt;
            var formats = new[] { "dd.MM.yyyy", "d.M.yyyy", "dd.MM.yy", "d.M.yy", "yyyy-MM-dd" };
            if (DateTime.TryParseExact(geb, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
            {
                return dt.ToString("yyMMdd");
            }
            var digits = new string(geb.Where(char.IsDigit).ToArray());
            if (digits.Length == 6) return digits;
            if (digits.Length == 8) return digits.Substring(2, 6);
            return "";
        }

        private static string BereinigenNachnameFuerMail(string nachname)
        {
            if (string.IsNullOrWhiteSpace(nachname)) return "";
            var n = UmlauteErsetzen(nachname.Trim().ToLowerInvariant());
            var cleaned = new string(n.Where(c => c >= 'a' && c <= 'z').ToArray());
            return string.IsNullOrEmpty(cleaned) ? "" : cleaned;
        }

        private static string BereinigenVornameFuerMail(string vorname)
        {
            if (string.IsNullOrWhiteSpace(vorname)) return "";
            var v = UmlauteErsetzen(vorname.Trim().ToLowerInvariant());
            var cleaned = new string(v.Where(c => c >= 'a' && c <= 'z').ToArray());
            return string.IsNullOrEmpty(cleaned) ? "" : cleaned;
        }

        private static string UmlauteErsetzen(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            return s
                .Replace("ä", "ae")
                .Replace("ö", "oe")
                .Replace("ü", "ue")
                .Replace("Ä", "Ae")
                .Replace("Ö", "Oe")
                .Replace("Ü", "Ue")
                .Replace("ß", "ss");
        }
    }
}
