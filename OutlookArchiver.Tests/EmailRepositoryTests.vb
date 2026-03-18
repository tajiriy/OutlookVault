Option Explicit On
Option Strict On
Option Infer Off

Imports NUnit.Framework
Imports OutlookArchiver.Data
Imports OutlookArchiver.Models
Imports System.IO

Namespace Tests

    <TestFixture>
    Public Class EmailRepositoryTests

        Private _testDbPath As String
        Private _dbManager As DatabaseManager
        Private _repo As EmailRepository

        <SetUp>
        Public Sub SetUp()
            _testDbPath = Path.Combine(Path.GetTempPath(), "OutlookArchiverTest_" & Guid.NewGuid().ToString("N") & ".db")
            _dbManager = New DatabaseManager(_testDbPath)
            _dbManager.Initialize()
            _repo = New EmailRepository(_dbManager)
        End Sub

        <TearDown>
        Public Sub TearDown()
            If File.Exists(_testDbPath) Then File.Delete(_testDbPath)
            If File.Exists(_testDbPath & "-wal") Then File.Delete(_testDbPath & "-wal")
            If File.Exists(_testDbPath & "-shm") Then File.Delete(_testDbPath & "-shm")
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  Email 挿入・取得
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub InsertEmail_ReturnsPositiveId()
            Dim email As Email = CreateTestEmail("test@example.com")

            Dim id As Integer = _repo.InsertEmail(email)

            Assert.IsTrue(id > 0)
        End Sub

        <Test>
        Public Sub GetEmailById_ReturnsInsertedEmail()
            Dim email As Email = CreateTestEmail("test@example.com")
            email.Subject = "テストメール"
            email.SenderName = "テスト太郎"
            email.SenderEmail = "taro@example.com"
            email.BodyText = "本文テキスト"
            email.FolderName = "受信トレイ"
            email.EmailSize = 12345

            Dim id As Integer = _repo.InsertEmail(email)
            Dim result As Email = _repo.GetEmailById(id)

            Assert.IsNotNull(result)
            Assert.AreEqual(id, result.Id)
            Assert.AreEqual("test@example.com", result.MessageId)
            Assert.AreEqual("テストメール", result.Subject)
            Assert.AreEqual("テスト太郎", result.SenderName)
            Assert.AreEqual("taro@example.com", result.SenderEmail)
            Assert.AreEqual("本文テキスト", result.BodyText)
            Assert.AreEqual("受信トレイ", result.FolderName)
            Assert.AreEqual(12345L, result.EmailSize)
        End Sub

        <Test>
        Public Sub GetEmailById_NonExistent_ReturnsNothing()
            Dim result As Email = _repo.GetEmailById(9999)

            Assert.IsNull(result)
        End Sub

        <Test>
        Public Sub GetEmailByMessageId_ReturnsCorrectEmail()
            Dim email As Email = CreateTestEmail("unique-msg@example.com")
            email.ThreadId = "thread-001"
            _repo.InsertEmail(email)

            Dim result As Email = _repo.GetEmailByMessageId("unique-msg@example.com")

            Assert.IsNotNull(result)
            Assert.AreEqual("unique-msg@example.com", result.MessageId)
            Assert.AreEqual("thread-001", result.ThreadId)
        End Sub

        <Test>
        Public Sub GetEmailByMessageId_NonExistent_ReturnsNothing()
            Dim result As Email = _repo.GetEmailByMessageId("nonexistent@example.com")

            Assert.IsNull(result)
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  メール一覧取得
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub GetEmails_ReturnsAllEmails()
            _repo.InsertEmail(CreateTestEmail("msg1@example.com"))
            _repo.InsertEmail(CreateTestEmail("msg2@example.com"))
            _repo.InsertEmail(CreateTestEmail("msg3@example.com"))

            Dim result As List(Of Email) = _repo.GetEmails()

            Assert.AreEqual(3, result.Count)
        End Sub

        <Test>
        Public Sub GetEmails_FilterByFolder_ReturnsOnlyMatchingFolder()
            Dim email1 As Email = CreateTestEmail("msg1@example.com")
            email1.FolderName = "受信トレイ"
            Dim email2 As Email = CreateTestEmail("msg2@example.com")
            email2.FolderName = "送信済み"

            _repo.InsertEmail(email1)
            _repo.InsertEmail(email2)

            Dim result As List(Of Email) = _repo.GetEmails(folderName:="受信トレイ")

            Assert.AreEqual(1, result.Count)
            Assert.AreEqual("受信トレイ", result(0).FolderName)
        End Sub

        <Test>
        Public Sub GetEmails_Paging_ReturnsCorrectPage()
            For i As Integer = 1 To 5
                Dim email As Email = CreateTestEmail("msg" & i.ToString() & "@example.com")
                email.ReceivedAt = New DateTime(2025, 1, i)
                _repo.InsertEmail(email)
            Next

            Dim page As List(Of Email) = _repo.GetEmails(pageIndex:=0, pageSize:=2)

            Assert.AreEqual(2, page.Count)
        End Sub

        <Test>
        Public Sub GetEmails_OrderedByReceivedAtDesc()
            Dim email1 As Email = CreateTestEmail("old@example.com")
            email1.ReceivedAt = New DateTime(2025, 1, 1)
            Dim email2 As Email = CreateTestEmail("new@example.com")
            email2.ReceivedAt = New DateTime(2025, 6, 1)

            _repo.InsertEmail(email1)
            _repo.InsertEmail(email2)

            Dim result As List(Of Email) = _repo.GetEmails()

            Assert.AreEqual("new@example.com", result(0).MessageId)
            Assert.AreEqual("old@example.com", result(1).MessageId)
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  フォルダ名一覧
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub GetFolderNames_ReturnsDistinctFolders()
            Dim email1 As Email = CreateTestEmail("msg1@example.com")
            email1.FolderName = "受信トレイ"
            Dim email2 As Email = CreateTestEmail("msg2@example.com")
            email2.FolderName = "送信済み"
            Dim email3 As Email = CreateTestEmail("msg3@example.com")
            email3.FolderName = "受信トレイ"

            _repo.InsertEmail(email1)
            _repo.InsertEmail(email2)
            _repo.InsertEmail(email3)

            Dim folders As List(Of String) = _repo.GetFolderNames()

            Assert.AreEqual(2, folders.Count)
            Assert.IsTrue(folders.Contains("受信トレイ"))
            Assert.IsTrue(folders.Contains("送信済み"))
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  総数
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub GetTotalCount_ReturnsCorrectCount()
            _repo.InsertEmail(CreateTestEmail("msg1@example.com"))
            _repo.InsertEmail(CreateTestEmail("msg2@example.com"))

            Assert.AreEqual(2, _repo.GetTotalCount())
        End Sub

        <Test>
        Public Sub GetTotalCount_WithFolder_ReturnsFilteredCount()
            Dim email1 As Email = CreateTestEmail("msg1@example.com")
            email1.FolderName = "受信トレイ"
            Dim email2 As Email = CreateTestEmail("msg2@example.com")
            email2.FolderName = "送信済み"

            _repo.InsertEmail(email1)
            _repo.InsertEmail(email2)

            Assert.AreEqual(1, _repo.GetTotalCount(folderName:="受信トレイ"))
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  スレッド取得
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub GetEmailsByThreadId_ReturnsSameThread()
            Dim email1 As Email = CreateTestEmail("msg1@example.com")
            email1.ThreadId = "thread-A"
            email1.ReceivedAt = New DateTime(2025, 1, 1)
            Dim email2 As Email = CreateTestEmail("msg2@example.com")
            email2.ThreadId = "thread-A"
            email2.ReceivedAt = New DateTime(2025, 1, 2)
            Dim email3 As Email = CreateTestEmail("msg3@example.com")
            email3.ThreadId = "thread-B"

            _repo.InsertEmail(email1)
            _repo.InsertEmail(email2)
            _repo.InsertEmail(email3)

            Dim result As List(Of Email) = _repo.GetEmailsByThreadId("thread-A")

            Assert.AreEqual(2, result.Count)
            ' 時系列昇順
            Assert.AreEqual("msg1@example.com", result(0).MessageId)
            Assert.AreEqual("msg2@example.com", result(1).MessageId)
        End Sub

        <Test>
        Public Sub GetEmailsByNormalizedSubject_ReturnsMatching()
            Dim email1 As Email = CreateTestEmail("msg1@example.com")
            email1.NormalizedSubject = "会議の件"
            Dim email2 As Email = CreateTestEmail("msg2@example.com")
            email2.NormalizedSubject = "会議の件"
            Dim email3 As Email = CreateTestEmail("msg3@example.com")
            email3.NormalizedSubject = "別の件"

            _repo.InsertEmail(email1)
            _repo.InsertEmail(email2)
            _repo.InsertEmail(email3)

            Dim result As List(Of Email) = _repo.GetEmailsByNormalizedSubject("会議の件", 10)

            Assert.AreEqual(2, result.Count)
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  Attachment 挿入・取得
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub InsertAttachment_ReturnsPositiveId()
            Dim emailId As Integer = _repo.InsertEmail(CreateTestEmail("msg@example.com"))
            Dim att As New Attachment()
            att.EmailId = emailId
            att.FileName = "document.pdf"
            att.FilePath = "2025\01\document.pdf"
            att.FileSize = 1024
            att.MimeType = "application/pdf"

            Dim attId As Integer = _repo.InsertAttachment(att)

            Assert.IsTrue(attId > 0)
        End Sub

        <Test>
        Public Sub InsertAttachment_InlineImage_SavesContentIdAndIsInline()
            Dim emailId As Integer = _repo.InsertEmail(CreateTestEmail("msg@example.com"))
            Dim att As New Attachment()
            att.EmailId = emailId
            att.FileName = "image001.png"
            att.FilePath = "2025\01\image001.png"
            att.FileSize = 2048
            att.MimeType = "image/png"
            att.ContentId = "image001@01D9ABCD.12345678"
            att.IsInline = True

            Dim attId As Integer = _repo.InsertAttachment(att)

            ' GetAttachmentsByEmailId は AppSettings.Instance.AttachmentDirectory を参照するため
            ' ここでは DB 直接クエリで確認
            Using conn As System.Data.SQLite.SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New System.Data.SQLite.SQLiteCommand(
                    "SELECT content_id, is_inline FROM attachments WHERE id = @id", conn)
                    cmd.Parameters.AddWithValue("@id", CType(attId, Object))
                    Using reader As System.Data.SQLite.SQLiteDataReader = cmd.ExecuteReader()
                        Assert.IsTrue(reader.Read())
                        Assert.AreEqual("image001@01D9ABCD.12345678", reader.GetString(0))
                        Assert.AreEqual(1, reader.GetInt32(1))
                    End Using
                End Using
            End Using
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  重複チェック
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub MessageIdExists_ExistingId_ReturnsTrue()
            _repo.InsertEmail(CreateTestEmail("existing@example.com"))

            Assert.IsTrue(_repo.MessageIdExists("existing@example.com"))
        End Sub

        <Test>
        Public Sub MessageIdExists_NonExistingId_ReturnsFalse()
            Assert.IsFalse(_repo.MessageIdExists("nonexistent@example.com"))
        End Sub

        <Test>
        Public Sub GetAllMessageIds_ReturnsAllIds()
            _repo.InsertEmail(CreateTestEmail("msg1@example.com"))
            _repo.InsertEmail(CreateTestEmail("msg2@example.com"))

            Dim ids As HashSet(Of String) = _repo.GetAllMessageIds()

            Assert.AreEqual(2, ids.Count)
            Assert.IsTrue(ids.Contains("msg1@example.com"))
            Assert.IsTrue(ids.Contains("msg2@example.com"))
        End Sub

        <Test>
        Public Sub GetAllMessageIds_CaseInsensitive()
            _repo.InsertEmail(CreateTestEmail("Test@Example.COM"))

            Dim ids As HashSet(Of String) = _repo.GetAllMessageIds()

            Assert.IsTrue(ids.Contains("test@example.com"))
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  削除・トゥームストーン
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub DeleteEmail_RemovesFromDb()
            Dim id As Integer = _repo.InsertEmail(CreateTestEmail("del@example.com"))

            _repo.DeleteEmail(id)

            Assert.IsNull(_repo.GetEmailById(id))
        End Sub

        <Test>
        Public Sub DeleteEmail_RecordsTombstone()
            Dim id As Integer = _repo.InsertEmail(CreateTestEmail("del@example.com"))

            _repo.DeleteEmail(id)

            Assert.IsTrue(_repo.IsMessageIdDeleted("del@example.com"))
        End Sub

        <Test>
        Public Sub IsMessageIdDeleted_NotDeleted_ReturnsFalse()
            Assert.IsFalse(_repo.IsMessageIdDeleted("notdeleted@example.com"))
        End Sub

        <Test>
        Public Sub GetAllDeletedMessageIds_ReturnsAllDeleted()
            _repo.MarkMessageIdAsDeleted("del1@example.com")
            _repo.MarkMessageIdAsDeleted("del2@example.com")

            Dim deleted As HashSet(Of String) = _repo.GetAllDeletedMessageIds()

            Assert.AreEqual(2, deleted.Count)
            Assert.IsTrue(deleted.Contains("del1@example.com"))
            Assert.IsTrue(deleted.Contains("del2@example.com"))
        End Sub

        <Test>
        Public Sub MarkMessageIdAsDeleted_Duplicate_NoError()
            _repo.MarkMessageIdAsDeleted("dup@example.com")

            Assert.DoesNotThrow(Sub() _repo.MarkMessageIdAsDeleted("dup@example.com"))
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  全文検索（FTS4）
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub SearchEmails_AsciiQuery_FindsBySubject()
            Dim email As Email = CreateTestEmail("search1@example.com")
            email.Subject = "Important Meeting"
            email.BodyText = "Please attend the meeting."
            _repo.InsertEmail(email)

            Dim results As List(Of Email) = _repo.SearchEmails("Meeting")

            Assert.AreEqual(1, results.Count)
            Assert.AreEqual("search1@example.com", results(0).MessageId)
        End Sub

        <Test>
        Public Sub SearchEmails_AsciiQuery_FindsBySenderName()
            Dim email As Email = CreateTestEmail("search2@example.com")
            email.SenderName = "John Smith"
            _repo.InsertEmail(email)

            Dim results As List(Of Email) = _repo.SearchEmails("Smith")

            Assert.AreEqual(1, results.Count)
        End Sub

        <Test>
        Public Sub SearchEmails_JapaneseQuery_FallsBackToLike()
            Dim email As Email = CreateTestEmail("search3@example.com")
            email.Subject = "週次ミーティング"
            email.BodyText = "来週の予定を確認してください。"
            _repo.InsertEmail(email)

            Dim results As List(Of Email) = _repo.SearchEmails("ミーティング")

            Assert.AreEqual(1, results.Count)
        End Sub

        <Test>
        Public Sub SearchEmails_JapaneseQuery_FindsByBodyText()
            Dim email As Email = CreateTestEmail("search4@example.com")
            email.Subject = "件名"
            email.BodyText = "この報告書を確認してください。"
            _repo.InsertEmail(email)

            Dim results As List(Of Email) = _repo.SearchEmails("報告書")

            Assert.AreEqual(1, results.Count)
        End Sub

        <Test>
        Public Sub SearchEmails_WithFolder_FiltersResults()
            Dim email1 As Email = CreateTestEmail("search5@example.com")
            email1.Subject = "Report"
            email1.FolderName = "受信トレイ"
            Dim email2 As Email = CreateTestEmail("search6@example.com")
            email2.Subject = "Report"
            email2.FolderName = "送信済み"

            _repo.InsertEmail(email1)
            _repo.InsertEmail(email2)

            Dim results As List(Of Email) = _repo.SearchEmails("Report", folderName:="受信トレイ")

            Assert.AreEqual(1, results.Count)
            Assert.AreEqual("search5@example.com", results(0).MessageId)
        End Sub

        <Test>
        Public Sub SearchEmails_NoMatch_ReturnsEmptyList()
            _repo.InsertEmail(CreateTestEmail("msg@example.com"))

            Dim results As List(Of Email) = _repo.SearchEmails("NONEXISTENT")

            Assert.AreEqual(0, results.Count)
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  バルクモード
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub BulkMode_IsBulkActive_ReturnsTrueWhileActive()
            Assert.IsFalse(_repo.IsBulkActive)

            _repo.BeginBulk()
            Assert.IsTrue(_repo.IsBulkActive)

            _repo.CommitBulk()
            Assert.IsFalse(_repo.IsBulkActive)
        End Sub

        <Test>
        Public Sub BulkMode_CommitPersistsData()
            _repo.BeginBulk()
            _repo.InsertEmail(CreateTestEmail("bulk1@example.com"))
            _repo.InsertEmail(CreateTestEmail("bulk2@example.com"))
            _repo.CommitBulk()

            Assert.AreEqual(2, _repo.GetTotalCount())
        End Sub

        <Test>
        Public Sub BulkMode_RollbackDiscardsData()
            _repo.BeginBulk()
            _repo.InsertEmail(CreateTestEmail("rollback1@example.com"))
            _repo.InsertEmail(CreateTestEmail("rollback2@example.com"))
            _repo.RollbackBulk()

            Assert.AreEqual(0, _repo.GetTotalCount())
        End Sub

        <Test>
        Public Sub BulkMode_BeginBulkTwice_NoError()
            _repo.BeginBulk()
            Assert.DoesNotThrow(Sub() _repo.BeginBulk())
            _repo.CommitBulk()
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  スレッド判定キャッシュ構築
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub GetThreadIdCaches_BuildsCorrectMaps()
            Dim email1 As Email = CreateTestEmail("msg1@example.com")
            email1.ThreadId = "thread-1"
            email1.NormalizedSubject = "会議"
            Dim email2 As Email = CreateTestEmail("msg2@example.com")
            email2.ThreadId = "thread-2"
            email2.NormalizedSubject = "報告"

            _repo.InsertEmail(email1)
            _repo.InsertEmail(email2)

            Dim msgIdMap As Dictionary(Of String, String) = Nothing
            Dim subjectMap As Dictionary(Of String, String) = Nothing
            _repo.GetThreadIdCaches(msgIdMap, subjectMap)

            Assert.AreEqual(2, msgIdMap.Count)
            Assert.AreEqual("thread-1", msgIdMap("msg1@example.com"))
            Assert.AreEqual("thread-2", msgIdMap("msg2@example.com"))

            Assert.AreEqual(2, subjectMap.Count)
            Assert.AreEqual("thread-1", subjectMap("会議"))
            Assert.AreEqual("thread-2", subjectMap("報告"))
        End Sub

        <Test>
        Public Sub GetThreadIdCaches_CaseInsensitiveKeys()
            Dim email As Email = CreateTestEmail("TEST@EXAMPLE.COM")
            email.ThreadId = "thread-1"
            email.NormalizedSubject = "Subject"
            _repo.InsertEmail(email)

            Dim msgIdMap As Dictionary(Of String, String) = Nothing
            Dim subjectMap As Dictionary(Of String, String) = Nothing
            _repo.GetThreadIdCaches(msgIdMap, subjectMap)

            Assert.IsTrue(msgIdMap.ContainsKey("test@example.com"))
            Assert.IsTrue(subjectMap.ContainsKey("subject"))
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  MessageID 重複挿入のエラー
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub InsertEmail_DuplicateMessageId_ThrowsException()
            _repo.InsertEmail(CreateTestEmail("dup@example.com"))

            Assert.Throws(Of System.Data.SQLite.SQLiteException)(
                Sub() _repo.InsertEmail(CreateTestEmail("dup@example.com")))
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  FTS トリガーの動作確認
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub FtsTrigger_DeleteEmail_RemovesFromFts()
            Dim email As Email = CreateTestEmail("fts-del@example.com")
            email.Subject = "UniqueSearchTerm"
            Dim id As Integer = _repo.InsertEmail(email)

            ' 削除前: 検索でヒットする
            Dim before As List(Of Email) = _repo.SearchEmails("UniqueSearchTerm")
            Assert.AreEqual(1, before.Count)

            _repo.DeleteEmail(id)

            ' 削除後: 検索でヒットしない
            Dim after As List(Of Email) = _repo.SearchEmails("UniqueSearchTerm")
            Assert.AreEqual(0, after.Count)
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  HasAttachments フラグ
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub InsertEmail_HasAttachmentsFlag_PersistsCorrectly()
            Dim email As Email = CreateTestEmail("att@example.com")
            email.HasAttachments = True

            Dim id As Integer = _repo.InsertEmail(email)
            Dim result As Email = _repo.GetEmailById(id)

            Assert.IsTrue(result.HasAttachments)
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  ヘルパー
        ' ════════════════════════════════════════════════════════════

        Private Shared Function CreateTestEmail(messageId As String) As Email
            Dim email As New Email()
            email.MessageId = messageId
            email.Subject = "Test Subject"
            email.NormalizedSubject = "Test Subject"
            email.SenderName = "Sender"
            email.SenderEmail = "sender@example.com"
            email.ReceivedAt = New DateTime(2025, 6, 1)
            email.FolderName = "受信トレイ"
            Return email
        End Function

    End Class

End Namespace
