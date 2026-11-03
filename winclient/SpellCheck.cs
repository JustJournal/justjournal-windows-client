using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using WeCantSpell.Hunspell;

namespace JustJournal
{
    /// <summary>
    ///     Spell checking with the Hunspell dictionaries in the dictionaries folder next to the program.
    /// </summary>
    public static class SpellCheck
    {
        private const string DefaultLanguage = "en_US";

        private static WordList _words;

        private static string DictionaryFolder
        {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "dictionaries"); }
        }

        /// <summary>
        ///     Load the dictionary for the user's language, or en_US when there isn't one.
        /// </summary>
        /// <returns>false when no dictionary was found</returns>
        public static bool Load()
        {
            if (_words != null)
                return true;

            foreach (string language in new[] {CultureInfo.CurrentUICulture.Name.Replace('-', '_'), DefaultLanguage})
            {
                string dic = Path.Combine(DictionaryFolder, language + ".dic");
                string aff = Path.Combine(DictionaryFolder, language + ".aff");
                if (File.Exists(dic) && File.Exists(aff))
                {
                    _words = WordList.CreateFromFiles(dic, aff);
                    return true;
                }
            }
            return false;
        }

        public static IEnumerable<string> Suggest(string word)
        {
            return _words.Suggest(word);
        }

        /// <summary>
        ///     Find the next misspelled word at or after start. Call Load first.
        /// </summary>
        /// <param name="ignored">words the user chose to ignore</param>
        public static bool FindMisspelling(string text, int start, ICollection<string> ignored,
                                           out int wordStart, out int wordLength)
        {
            int cur = start;
            while (cur < text.Length)
            {
                while (cur < text.Length && !IsWordChar(text[cur]))
                    cur++;
                int begin = cur;
                bool hasDigit = false;
                while (cur < text.Length && IsWordChar(text[cur]))
                {
                    if (char.IsDigit(text[cur]))
                        hasDigit = true;
                    cur++;
                }

                // quotes around a word aren't part of it
                int end = cur;
                while (begin < end && IsApostrophe(text[begin]))
                    begin++;
                while (end > begin && IsApostrophe(text[end - 1]))
                    end--;

                // numbers and words like 2nd are left alone
                if (end > begin && !hasDigit)
                {
                    string word = text.Substring(begin, end - begin);
                    if (!ignored.Contains(word) && !_words.Check(word))
                    {
                        wordStart = begin;
                        wordLength = end - begin;
                        return true;
                    }
                }
            }

            wordStart = 0;
            wordLength = 0;
            return false;
        }

        private static bool IsWordChar(char c)
        {
            return char.IsLetterOrDigit(c) || IsApostrophe(c);
        }

        private static bool IsApostrophe(char c)
        {
            return c == '\'' || c == '’';
        }
    }

    public enum SpellingChoice
    {
        Change,
        Ignore,
        IgnoreAll,
        Cancel
    }

    /// <summary>
    ///     Asks what to do with one misspelled word.
    /// </summary>
    public class SpellingDialog : Form
    {
        private readonly TextBox _replacement = new TextBox();
        private readonly ListBox _suggestions = new ListBox();

        public SpellingDialog(string word, IEnumerable<string> suggestions)
        {
            Text = @"Check Spelling";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(360, 200);
            Choice = SpellingChoice.Cancel;

            var label = new Label {Text = @"Not in dictionary: " + word, AutoEllipsis = true};
            label.SetBounds(12, 12, 250, 20);
            _replacement.SetBounds(12, 36, 250, 20);
            _replacement.Text = word;
            _suggestions.SetBounds(12, 62, 250, 126);
            _suggestions.SelectedIndexChanged += delegate
                {
                    if (_suggestions.SelectedItem != null)
                        _replacement.Text = (string) _suggestions.SelectedItem;
                };
            _suggestions.DoubleClick += delegate
                {
                    Choice = SpellingChoice.Change;
                    DialogResult = DialogResult.OK;
                };
            foreach (string suggestion in suggestions)
                _suggestions.Items.Add(suggestion);
            if (_suggestions.Items.Count > 0)
                _suggestions.SelectedIndex = 0;

            Controls.AddRange(new Control[] {label, _replacement, _suggestions});
            AcceptButton = AddButton(@"&Change", 12, SpellingChoice.Change);
            AddButton(@"&Ignore", 42, SpellingChoice.Ignore);
            AddButton(@"Ignore &All", 72, SpellingChoice.IgnoreAll);
            CancelButton = AddButton(@"Cancel", 165, SpellingChoice.Cancel);
        }

        public SpellingChoice Choice { get; private set; }

        public string Replacement
        {
            get { return _replacement.Text; }
        }

        private Button AddButton(string text, int top, SpellingChoice choice)
        {
            var button = new Button {Text = text};
            button.SetBounds(273, top, 75, 23);
            button.Click += delegate
                {
                    Choice = choice;
                    DialogResult = DialogResult.OK;
                };
            Controls.Add(button);
            return button;
        }
    }
}
