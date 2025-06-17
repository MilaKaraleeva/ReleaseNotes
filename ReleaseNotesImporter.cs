using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xceed.Words.NET;
namespace ReleaseNotesImporter
{
   public class ReleaseNote
   {
       public DateTime Date { get ; set; }
       public string TicketId { get; set; }
       public string Description { get; set; }
   }
   public class ReleaseNoteImporter
   {
       public List<ReleaseNote> ImportFromFolder(string folderPath)
       {
           var allNotes = new List<ReleaseNote>();
           var docFiles = Directory.GetFiles(folderPath, "*.docx");
           foreach (var file in docFiles)
           {
               var notes = ParseWordFile(file);
               allNotes.AddRange(notes);
           }
           return allNotes;
       }
       private List<ReleaseNote> ParseWordFile(string filePath)
       {
           var notes = new List<ReleaseNote>();
           using (var document = DocX.Load(filePath))
           {
                // Word document has paragraphs formatted like:
               // Date: 2025-06-10
               // Ticket: ANYPAY-1234
               // Title: Login bug in payment screen.
               var paragraphs = document.Paragraphs;
               DateTime currentDate = DateTime.MinValue;
               string currentTicket = string.Empty;
               string currentDescription = string.Empty;
               foreach (var para in paragraphs)
               {
                   var text = para.Text.Trim();
                   if (text.StartsWith("Date:", StringComparison.OrdinalIgnoreCase))
                   {
                       DateTime.TryParse(text.Substring(5).Trim(), out currentDate);
                   }
                   else if (text.StartsWith("Ticket:", StringComparison.OrdinalIgnoreCase))
                   {
                       currentTicket = text.Substring(7).Trim();
                   }
                   else if (text.StartsWith("Description:", StringComparison.OrdinalIgnoreCase))
                   {
                       currentDescription = text.Substring(12).Trim();
                       if (!string.IsNullOrEmpty(currentTicket) && currentDate != DateTime.MinValue)
                       {
                           notes.Add(new ReleaseNote
                           {
                               Date = currentDate,
                               TicketId = currentTicket,
                               Description = currentDescription
                           });
                           // Reset next ticket
                           currentTicket = string.Empty;
                           currentDescription = string.Empty;
                       }
                   }
               }
           }
           return notes;
       }
       public Dictionary<DateTime, List<ReleaseNote>> GroupByDate(List<ReleaseNote> notes)
       {
           return notes
               .GroupBy(n => n.Date.Date)
               .ToDictionary(g => g.Key, g => g.ToList());
       }
   }}