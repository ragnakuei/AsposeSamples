﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;

namespace AsposeWordSamples;

class Program
{
    protected static void Main(string[] args)
    {
        ReplaceSample();
    }

    private static void ReplaceSample()
    {
        // 副檔名一定要正確，否則會判斷失敗

        var templateFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "docx", "ReplaceSample.docx");
        var outputFile   = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "docx", "AfterReplaceSample.docx");

        if (File.Exists(outputFile))
        {
            File.Delete(outputFile);
        }

        var doc = new Document(templateFile);

        ReplaceTexts(doc);

        MergeFields(doc);

        Table1(doc);

        doc.Save(outputFile);
    }

    /// <summary>
    /// Replace 方式
    /// </summary>
    /// <param name="doc"></param>
    private static void ReplaceTexts(Document doc)
    {
        var replaceFields = new Dictionary<string, string>
        {
            ["_fullName_"] = "Test",
            ["_age_"]      = "18",
        };

        foreach (var field in replaceFields)
        {
            // 新版語法才支援 FindReplaceOptions
            doc.Range.Replace(field.Key, field.Value, new FindReplaceOptions
            {
                Direction          = FindReplaceDirection.Forward,
                FindWholeWordsOnly = false,
                IgnoreDeleted      = false,
                IgnoreFields       = false,
                IgnoreInserted     = false,
                LegacyMode         = false,
                MatchCase          = true,
                ReplacingCallback  = null,
                UseLegacyOrder     = false,
                UseSubstitutions   = false
            });

            // 舊版語法
            // doc.Range.Replace(field.Key, field.Value, isMatchCase: true, isMatchWholeWord: true);
        }
    }

    /// <summary>
    /// MergeField 方式
    /// </summary>
    private static void MergeFields(Document doc)
    {
        var mergeFields = new Dictionary<string, string>
        {
            ["Column1"] = "A",
            ["Column2"] = "B",
        };

        doc.MailMerge.Execute(mergeFields.Keys.ToArray(), mergeFields.Values.ToArray());
    }

    /// <summary>
    /// 處理 Table
    /// </summary>
    /// <remarks>
    /// 1. 透過 Bookmark 取得 Table
    /// 2. 範本內的表格，都預先輸入一個字，並且設定好 Style，之後就只需要直接替換其文字就可以了 !
    /// 3. 資料列的部份，預先留好一列，之後就可以複製該列，來新增資料列
    /// </remarks>
    private static void Table1(Document doc)
    {
        // 從 Bookmark Table1 來取得 Table
        var table = doc.Range.Bookmarks["Table1"].BookmarkStart.GetAncestor(NodeType.Table) as Table;

        // 處理標題列
        table.Rows[0].Cells[1].FirstParagraph.Runs[0].Text = DateTime.Now.ToString("yyyyMMdd");
        table.Rows[0].Cells[2].FirstParagraph.Runs[0].Text = DateTime.Now.AddYears(-1).ToString("yyyy");
        table.Rows[0].Cells[3].FirstParagraph.Runs[0].Text = DateTime.Now.AddYears(-2).ToString("yyyy");

        // 複製第二列，之後依照需求新增
        var row = table.Rows[1].Clone(true) as Row;

        // 刪除既有的第二列
        table.Rows.RemoveAt(1);

        var rows = new[]
        {
            new TableRowDto { Title = "Title1", Column1 = "1", Column2 = "2", Column3 = "3" },
            new TableRowDto { Title = "Title2", Column1 = "4", Column2 = "5", Column3 = "6" },
            new TableRowDto { Title = "Titls3", Column1 = "7", Column2 = "8", Column3 = "9" },
        };

        foreach (var item in rows)
        {
            var newRow = row.Clone(true) as Row;

            newRow.Cells[0].FirstParagraph.Runs[0].Text = item.Title;
            newRow.Cells[1].FirstParagraph.Runs[0].Text = item.Column1;
            newRow.Cells[2].FirstParagraph.Runs[0].Text = item.Column2;
            newRow.Cells[3].FirstParagraph.Runs[0].Text = item.Column3;

            table.Rows.Add(newRow);
        }
    }

    private class TableRowDto
    {
        public string Title { get; set; }

        public string? Column1 { get; set; }

        public string? Column2 { get; set; }

        public string? Column3 { get; set; }
    }
}