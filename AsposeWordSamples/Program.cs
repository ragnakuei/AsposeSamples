using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Replacing;

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

        // Replace 方式
        var replaceFields = new Dictionary<string, string>
        {
            ["_fullName_"] = "Test",
            ["_age_"]      = "18",
        };

        var doc = new Document(templateFile);

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

        // MergeField 方式
        var mergeFields = new Dictionary<string, string>
        {
            ["Column1"] = "A",
            ["Column2"] = "B",
        };
        doc.MailMerge.Execute(mergeFields.Keys.ToArray(), mergeFields.Values.ToArray());
        
        doc.Save(outputFile);
    }
}