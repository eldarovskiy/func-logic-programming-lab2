using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Microsoft.Win32;

namespace LetterGeneratorApp;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private readonly LetterTemplateService templateService = new();
    private readonly List<LetterAppendix> appendices = new();

    public MainWindow()
    {
        InitializeComponent();
        LetterDatePicker.SelectedDate = DateTime.Today;
        RefreshAppendicesListBox();
    }

    private void AddAppendixButton_Click(object sender, RoutedEventArgs e)
    {
        string title = AppendixTitleTextBox.Text.Trim();
        string body = AppendixBodyTextBox.Text.Trim();

        if (string.IsNullOrWhiteSpace(title) || string.IsNullOrWhiteSpace(body))
        {
            StatusTextBlock.Text = "Для приложения заполните заголовок и текст.";
            return;
        }

        int? pageCount = null;
        string pagesRaw = AppendixPagesTextBox.Text.Trim();
        if (!string.IsNullOrWhiteSpace(pagesRaw) && int.TryParse(pagesRaw, out int parsed) && parsed > 0)
        {
            pageCount = parsed;
        }

        appendices.Add(new LetterAppendix(title, body, pageCount));
        RefreshAppendicesListBox();

        AppendixTitleTextBox.Text = string.Empty;
        AppendixBodyTextBox.Text = string.Empty;
        AppendixPagesTextBox.Text = string.Empty;

        StatusTextBlock.Text = $"Добавлено приложений: {appendices.Count}.";
    }

    private void ClearAppendicesButton_Click(object sender, RoutedEventArgs e)
    {
        appendices.Clear();
        RefreshAppendicesListBox();
        StatusTextBlock.Text = "Список приложений очищен.";
    }

    private void RefreshAppendicesListBox()
    {
        if (AppendicesListBox == null)
        {
            return;
        }

        AppendicesListBox.ItemsSource = null;
        AppendicesListBox.ItemsSource = appendices
            .Select((x, i) => x.ToListLine(i + 1))
            .ToList();
    }

    private void CreateDocumentButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            string templatePath = templateService.EnsureTemplateExists();

            SaveFileDialog saveFileDialog = new()
            {
                Filter = "Документ Word (*.docx)|*.docx",
                FileName = $"Письмо_{DateTime.Now:yyyy-MM-dd_HH-mm}.docx",
                InitialDirectory = Path.Combine(AppContext.BaseDirectory, "Generated")
            };

            if (saveFileDialog.ShowDialog() != true)
            {
                StatusTextBlock.Text = "Сохранение отменено.";
                return;
            }

            Dictionary<string, string> replacements = new()
            {
                ["{RECIPIENT_POST}"] = RecipientPostTextBox.Text.Trim(),
                ["{RECIPIENT_NAME}"] = RecipientNameTextBox.Text.Trim(),
                ["{RECIPIENT_GREETING}"] = RecipientGreetingTextBox.Text.Trim(),
                ["{SENDER_POST}"] = SenderPostTextBox.Text.Trim(),
                ["{SENDER_NAME}"] = SenderNameTextBox.Text.Trim(),
                ["{LETTER_SUBJECT}"] = SubjectTextBox.Text.Trim(),
                ["{LETTER_BODY}"] = LetterBodyTextBox.Text.Trim(),
                ["{LETTER_DATE}"] = (LetterDatePicker.SelectedDate ?? DateTime.Today).ToString("dd.MM.yyyy")
            };

            List<LetterAppendix> appendicesForDocument = new(appendices);

            templateService.GenerateDocument(templatePath, saveFileDialog.FileName, replacements, appendicesForDocument);

            StatusTextBlock.Text = $"Документ создан: {saveFileDialog.FileName}";

            Process.Start(new ProcessStartInfo(saveFileDialog.FileName)
            {
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"Ошибка: {ex.Message}";
            MessageBox.Show(this, ex.Message, "Не удалось создать документ", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}

internal sealed class LetterTemplateService
{
    private static readonly string TemplatesDirectory = Path.Combine(AppContext.BaseDirectory, "Templates");
    private static readonly string TemplatePath = Path.Combine(TemplatesDirectory, "Шаблон.docx");

    public string EnsureTemplateExists()
    {
        Directory.CreateDirectory(TemplatesDirectory);

        CreateDefaultTemplate(TemplatePath);

        return TemplatePath;
    }

    public void GenerateDocument(
        string templatePath,
        string outputPath,
        IReadOnlyDictionary<string, string> replacements,
        IReadOnlyList<LetterAppendix> appendices)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? AppContext.BaseDirectory);
        File.Copy(templatePath, outputPath, overwrite: true);

        using WordprocessingDocument document = WordprocessingDocument.Open(outputPath, true);

        ReplaceInElement(document.MainDocumentPart?.Document, replacements);
        InsertAppendices(document.MainDocumentPart?.Document?.Body, appendices);

        if (document.MainDocumentPart != null)
        {
            foreach (HeaderPart headerPart in document.MainDocumentPart.HeaderParts)
            {
                ReplaceInElement(headerPart.Header, replacements);
            }

            foreach (FooterPart footerPart in document.MainDocumentPart.FooterParts)
            {
                ReplaceInElement(footerPart.Footer, replacements);
            }
        }

        document.MainDocumentPart?.Document.Save();
    }

    private static void InsertAppendices(Body? body, IReadOnlyList<LetterAppendix> appendices)
    {
        if (body == null)
        {
            return;
        }

        ReplacePlaceholderWithParagraphs(body, "{APPENDIX_LIST_PLACEHOLDER}", BuildAppendixListParagraphs(appendices));
        ReplacePlaceholderWithParagraphs(body, "{APPENDIX_CONTENT_PLACEHOLDER}", BuildAppendixContentParagraphs(appendices));
    }

    private static void ReplacePlaceholderWithParagraphs(Body body, string placeholder, IReadOnlyList<Paragraph> paragraphs)
    {
        Paragraph? marker = body.Elements<Paragraph>()
            .FirstOrDefault(p => p.InnerText.Contains(placeholder, StringComparison.Ordinal));

        if (marker == null)
        {
            return;
        }

        OpenXmlElement anchor = marker;
        foreach (Paragraph paragraph in paragraphs)
        {
            anchor.InsertAfterSelf(paragraph);
            anchor = paragraph;
        }

        marker.Remove();
    }

    private static IReadOnlyList<Paragraph> BuildAppendixListParagraphs(IReadOnlyList<LetterAppendix> appendices)
    {
        if (appendices.Count == 0)
        {
            return Array.Empty<Paragraph>();
        }

        List<Paragraph> result = new();
        string heading = appendices.Count == 1 ? "Приложение:" : "Приложения:";
        result.Add(CreateParagraph(heading, JustificationValues.Left, 18, true));

        for (int i = 0; i < appendices.Count; i++)
        {
            result.Add(CreateParagraph(appendices[i].ToListLine(i + 1), JustificationValues.Left, 18));
        }

        result.Add(CreateParagraph(string.Empty, JustificationValues.Left, 12));
        return result;
    }

    private static IReadOnlyList<Paragraph> BuildAppendixContentParagraphs(IReadOnlyList<LetterAppendix> appendices)
    {
        if (appendices.Count == 0)
        {
            return Array.Empty<Paragraph>();
        }

        List<Paragraph> result = new();

        for (int i = 0; i < appendices.Count; i++)
        {
            LetterAppendix appendix = appendices[i];
            string label = appendices.Count == 1 ? "Приложение" : $"Приложение {i + 1}";

            result.Add(CreateParagraph(label, JustificationValues.Right, 18, true));
            result.Add(CreateParagraph(appendix.Title, JustificationValues.Center, 20, true));
            result.Add(CreateParagraph(appendix.Body, JustificationValues.Both, 18));
            result.Add(CreateParagraph(string.Empty, JustificationValues.Left, 12));
        }

        return result;
    }

    private static void ReplaceInElement(OpenXmlElement? root, IReadOnlyDictionary<string, string> replacements)
    {
        if (root == null)
        {
            return;
        }

        foreach (Text text in root.Descendants<Text>())
        {
            string value = text.Text;

            foreach (KeyValuePair<string, string> replacement in replacements)
            {
                if (value.Contains(replacement.Key, StringComparison.Ordinal))
                {
                    value = value.Replace(replacement.Key, replacement.Value ?? string.Empty, StringComparison.Ordinal);
                }
            }

            text.Text = value;
        }
    }

    private static void CreateDefaultTemplate(string templatePath)
    {
        using WordprocessingDocument document = WordprocessingDocument.Create(templatePath, WordprocessingDocumentType.Document);

        MainDocumentPart mainDocumentPart = document.AddMainDocumentPart();
        mainDocumentPart.Document = new Document(new Body());

        Body body = mainDocumentPart.Document.Body!;
        body.Append(new SectionProperties(
            new PageSize() { Width = 12240U, Height = 15840U },
            new PageMargin() { Top = 900, Right = 900, Bottom = 900, Left = 900, Header = 450U, Footer = 450U, Gutter = 0U }));
        
        // Логотип вверху по центру (узкий)
        body.Append(CreateCenteredLogoBlock(mainDocumentPart));
        body.Append(CreateParagraph(string.Empty, JustificationValues.Center, 12));

        // Трёхколоночный блок: реквизиты слева, пусто в центре, адресат справа
        Table topTable = new(
            new TableProperties(new TableBorders(
                new TopBorder() { Val = BorderValues.Nil },
                new BottomBorder() { Val = BorderValues.Nil },
                new LeftBorder() { Val = BorderValues.Nil },
                new RightBorder() { Val = BorderValues.Nil },
                new InsideHorizontalBorder() { Val = BorderValues.Nil },
                new InsideVerticalBorder() { Val = BorderValues.Nil })),
            new TableGrid(new GridColumn() { Width = "3200" }, new GridColumn() { Width = "3200" }, new GridColumn() { Width = "3200" }));

        TableRow topRow = new();

        // Левая колонка: реквизиты банка
        TableCell left = new(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "3200" }));
        left.Append(CreateParagraph("Акционерный коммерческий банк", JustificationValues.Left, 22, true));
        left.Append(CreateParagraph("\"ИНЕКСБАНК\"", JustificationValues.Left, 20, true));
        left.Append(CreateParagraph("ул. Лубянка, 10, стр. 2", JustificationValues.Left, 18));
        left.Append(CreateParagraph("Москва, Россия, 120440", JustificationValues.Left, 18));
        left.Append(CreateParagraph("тел. (095) 228 09 78", JustificationValues.Left, 18));

        // Центр: пустое место
        TableCell center = new(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "3200" }));
        center.Append(CreateParagraph(string.Empty, JustificationValues.Center, 12));

        // Правая колонка: адресат и отправитель на одном уровне с реквизитами слева
        TableCell right = new(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "3200" }));
        right.Append(CreateParagraph("Кому:", JustificationValues.Right, 18, true));
        right.Append(CreateParagraph("{RECIPIENT_POST}", JustificationValues.Right, 16));
        right.Append(CreateParagraph("{RECIPIENT_NAME}", JustificationValues.Right, 16));
        right.Append(CreateParagraph(string.Empty, JustificationValues.Right, 16));
        right.Append(CreateParagraph("От:", JustificationValues.Right, 18, true));
        right.Append(CreateParagraph("{SENDER_POST}", JustificationValues.Right, 16));
        right.Append(CreateParagraph("{SENDER_NAME}", JustificationValues.Right, 16));

        topRow.Append(left);
        topRow.Append(center);
        topRow.Append(right);
        topTable.Append(topRow);
        body.Append(topTable);

        // Строки с датой и номером (слева) и пустой справа для баланса
        Table meta = new(
            new TableProperties(new TableBorders(
                new TopBorder() { Val = BorderValues.Nil },
                new BottomBorder() { Val = BorderValues.Nil },
                new LeftBorder() { Val = BorderValues.Nil },
                new RightBorder() { Val = BorderValues.Nil },
                new InsideHorizontalBorder() { Val = BorderValues.Nil },
                new InsideVerticalBorder() { Val = BorderValues.Nil })),
            new TableGrid(new GridColumn() { Width = "4800" }, new GridColumn() { Width = "4800" }));
        TableRow metaRow = new();
        metaRow.Append(CreateCell(new[] { CreateParagraph("{LETTER_DATE}    № 1 - 5 /29", JustificationValues.Left, 16) }, "4800"));
        metaRow.Append(CreateCell(new[] { CreateParagraph(string.Empty, JustificationValues.Left, 16) }, "4800"));
        meta.Append(metaRow);
        body.Append(meta);

        // Заголовок и центрированное обращение
        body.Append(CreateParagraph(string.Empty, JustificationValues.Center, 12));
        body.Append(CreateParagraph("Тема: {LETTER_SUBJECT}", JustificationValues.Left, 20, true));
        body.Append(CreateParagraph(string.Empty, JustificationValues.Center, 12));
        body.Append(CreateParagraph("Уважаемый {RECIPIENT_GREETING}!", JustificationValues.Center, 22, true));
        body.Append(CreateParagraph(string.Empty, JustificationValues.Left, 12));

        // Основной текст — по ширине страницы, но выровнен по ширине (Justify)
        body.Append(CreateParagraph("{LETTER_BODY}", JustificationValues.Both, 22));
        body.Append(CreateParagraph(string.Empty, JustificationValues.Left, 20));

        // Плейсхолдеры для автоматической вставки списка и содержимого приложений
        body.Append(CreateParagraph("{APPENDIX_LIST_PLACEHOLDER}", JustificationValues.Left, 18));
        body.Append(CreateParagraph("{APPENDIX_CONTENT_PLACEHOLDER}", JustificationValues.Left, 18));

        // Подписи: левый блок должность и контакт, справа подпись и ФИО
        Table signTable = new(
            new TableProperties(new TableBorders(
                new TopBorder() { Val = BorderValues.Nil },
                new BottomBorder() { Val = BorderValues.Nil },
                new LeftBorder() { Val = BorderValues.Nil },
                new RightBorder() { Val = BorderValues.Nil },
                new InsideHorizontalBorder() { Val = BorderValues.Nil },
                new InsideVerticalBorder() { Val = BorderValues.Nil })),
            new TableGrid(new GridColumn() { Width = "4800" }, new GridColumn() { Width = "4800" }));
        TableRow signRow = new();
        signRow.Append(CreateCell(new[] { CreateParagraph(string.Empty, JustificationValues.Left, 18), CreateParagraph(string.Empty, JustificationValues.Left, 18) }, "4800"));
        signRow.Append(CreateCell(new[] { CreateParagraph("Подпись ____________________", JustificationValues.Right, 18), CreateParagraph(string.Empty, JustificationValues.Right, 18) }, "4800"));
        signTable.Append(signRow);
        body.Append(signTable);

        // Футер
        body.Append(CreateParagraph(string.Empty, JustificationValues.Left, 12));
        body.Append(CreateParagraph("Автоматически сгенерировано приложением WPF", JustificationValues.Center, 16));

        mainDocumentPart.Document.Save();
    }

    private static Paragraph CreateCenteredLogoBlock(MainDocumentPart mainDocumentPart)
    {
        string logoPath = Path.Combine(TemplatesDirectory, "logo.png");

        if (!File.Exists(logoPath))
        {
            return new Paragraph(
                new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center },
                    new SpacingBetweenLines() { After = "40" }),
                new Run(
                    new RunProperties(new Bold(), new FontSize() { Val = "28" }),
                    new Text("[ ЛОГОТИП ]") { Space = SpaceProcessingModeValues.Preserve }));
        }

        ImagePart imagePart = mainDocumentPart.AddImagePart(ImagePartType.Png);
        using (FileStream stream = File.OpenRead(logoPath))
        {
            imagePart.FeedData(stream);
        }

        string relationshipId = mainDocumentPart.GetIdOfPart(imagePart);

        return new Paragraph(
            new ParagraphProperties(
                new Justification() { Val = JustificationValues.Center },
                new SpacingBetweenLines() { After = "40" }),
            new Run(CreateLogoDrawing(relationshipId)));
    }

    private static Drawing CreateLogoDrawing(string relationshipId)
    {
        return new Drawing(
            new DW.Inline(
                new DW.Extent() { Cx = 1300000L, Cy = 650000L },
                new DW.EffectExtent()
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DW.DocProperties() { Id = 1U, Name = "Logo" },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks() { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties() { Id = 0U, Name = "logo.png" },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip() { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset() { X = 0L, Y = 0L },
                                    new A.Extents() { Cx = 1300000L, Cy = 650000L }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
                    )
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                    })
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            });
    }

    private static Table CreateRecipientBlock()
    {
        Table table = new(
            new TableProperties(
                new TableBorders(
                    new TopBorder() { Val = BorderValues.Nil },
                    new BottomBorder() { Val = BorderValues.Nil },
                    new LeftBorder() { Val = BorderValues.Nil },
                    new RightBorder() { Val = BorderValues.Nil },
                    new InsideHorizontalBorder() { Val = BorderValues.Nil },
                    new InsideVerticalBorder() { Val = BorderValues.Nil })),
            new TableGrid(
                new GridColumn() { Width = "5200" },
                new GridColumn() { Width = "4400" }));

        TableRow row = new();
        row.Append(CreateCell(new[]
        {
            CreateParagraph("От 07.05.2026 № _____", JustificationValues.Left, 20),
            CreateParagraph("На № _____ от ________", JustificationValues.Left, 20)
        }, "5200"));

        row.Append(CreateCell(new[]
        {
            CreateParagraph("Генеральному директору", JustificationValues.Left, 20, true),
            CreateParagraph("ООО \"Художественная школа имени Репина\"", JustificationValues.Left, 20),
            CreateParagraph("Иванову Ивану Ивановичу", JustificationValues.Left, 20)
        }, "4400"));

        table.Append(row);
        return table;
    }

    private static TableCell CreateCell(IEnumerable<Paragraph> paragraphs, string width)
    {
        TableCell cell = new(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = width }));

        foreach (Paragraph paragraph in paragraphs)
        {
            cell.Append(paragraph);
        }

        return cell;
    }

    private static Paragraph CreateParagraph(string text, JustificationValues justification, int fontSizeHalfPoints, bool bold = false)
    {
        Paragraph paragraph = new(
            new ParagraphProperties(
                new Justification() { Val = justification },
                new SpacingBetweenLines() { Before = "0", After = "80" }));

        RunProperties runProperties = new();

        if (bold)
        {
            runProperties.Append(new Bold());
        }

        runProperties.Append(new FontSize() { Val = fontSizeHalfPoints.ToString() });

        paragraph.Append(
            new Run(
                runProperties,
                new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

        return paragraph;
    }

    private static Paragraph CreateParagraph(string text)
    {
        return new Paragraph(
            new ParagraphProperties(new SpacingBetweenLines() { After = "160" }),
            new Run(
                new RunProperties(new FontSize() { Val = "24" }),
                new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }
}

internal sealed class LetterAppendix
{
    public LetterAppendix(string title, string body, int? pageCount)
    {
        Title = title;
        Body = body;
        PageCount = pageCount;
    }

    public string Title { get; }

    public string Body { get; }

    public int? PageCount { get; }

    public string ToListLine(int index)
    {
        return PageCount.HasValue
            ? $"{index}. {Title} на {PageCount.Value} л."
            : $"{index}. {Title}";
    }
}