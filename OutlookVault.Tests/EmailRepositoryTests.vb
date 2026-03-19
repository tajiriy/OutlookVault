Option Explicit On
Option Strict On
Option Infer Off

Imports NUnit.Framework
Imports OutlookVault.Data
Imports OutlookVault.Models
Imports System.IO

Namespace Tests

    <TestFixture>
    Public Class EmailRepositoryTests

        Private _testDbPath As String
        Private _dbManager As DatabaseManager
        Private _repo As EmailRepository

        <SetUp>
        Public Sub SetUp()
            _testDbPath = Path.Combine(Path.GetTempPath(), "OutlookVaultTest_" & Guid.NewGuid().ToString("N") & ".db")
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
        Public Sub GetAllMessageIdFolderPairs_ReturnsAllPairs()
            _repo.InsertEmail(CreateTestEmail("msg1@example.com"))
            _repo.InsertEmail(CreateTestEmail("msg2@example.com"))

            Dim pairs As HashSet(Of String) = _repo.GetAllMessageIdFolderPairs()

            Assert.AreEqual(2, pairs.Count)
            Assert.IsTrue(pairs.Contains("msg1@example.com" & vbTab & "受信トレイ"))
            Assert.IsTrue(pairs.Contains("msg2@example.com" & vbTab & "受信トレイ"))
        End Sub

        <Test>
        Public Sub GetAllMessageIdFolderPairs_CaseInsensitive()
            _repo.InsertEmail(CreateTestEmail("Test@Example.COM"))

            Dim pairs As HashSet(Of String) = _repo.GetAllMessageIdFolderPairs()

            Assert.IsTrue(pairs.Contains("test@example.com" & vbTab & "受信トレイ"))
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  論理削除（ゴミ箱）
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub SoftDeleteEmail_MovesToTrash()
            Dim id As Integer = _repo.InsertEmail(CreateTestEmail("del@example.com"))

            _repo.SoftDeleteEmail(id)

            ' GetEmailById は論理削除でも取得できる（SELECT * に deleted_at フィルタなし）
            Dim email As Email = _repo.GetEmailById(id)
            Assert.IsNotNull(email)
            ' ただし一覧からは除外される
            Dim list As List(Of Email) = _repo.GetEmailsForList()
            Assert.IsFalse(list.Exists(Function(e) e.Id = id))
            ' ゴミ箱には表示される
            Dim trash As List(Of Email) = _repo.GetTrashEmails()
            Assert.IsTrue(trash.Exists(Function(e) e.Id = id))
        End Sub

        <Test>
        Public Sub SoftDeleteEmailsByIds_MovesMultipleToTrash()
            Dim id1 As Integer = _repo.InsertEmail(CreateTestEmail("del1@example.com"))
            Dim id2 As Integer = _repo.InsertEmail(CreateTestEmail("del2@example.com"))
            Dim id3 As Integer = _repo.InsertEmail(CreateTestEmail("keep@example.com"))

            _repo.SoftDeleteEmailsByIds(New List(Of Integer)() From {id1, id2})

            Assert.AreEqual(1, _repo.GetEmailsForList().Count)
            Assert.AreEqual(2, _repo.GetTrashCount())
        End Sub

        <Test>
        Public Sub RestoreEmailsByIds_RestoresFromTrash()
            Dim id As Integer = _repo.InsertEmail(CreateTestEmail("restore@example.com"))
            _repo.SoftDeleteEmailsByIds(New List(Of Integer)() From {id})
            Assert.AreEqual(1, _repo.GetTrashCount())

            _repo.RestoreEmailsByIds(New List(Of Integer)() From {id})

            Assert.AreEqual(0, _repo.GetTrashCount())
            Assert.AreEqual(1, _repo.GetEmailsForList().Count)
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  物理削除（パージ）・トゥームストーン
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub PurgeEmailsByIds_RemovesFromDb()
            Dim id As Integer = _repo.InsertEmail(CreateTestEmail("del@example.com"))

            _repo.PurgeEmailsByIds(New List(Of Integer)() From {id})

            Assert.IsNull(_repo.GetEmailById(id))
        End Sub

        <Test>
        Public Sub PurgeEmailsByIds_RecordsTombstone()
            Dim id As Integer = _repo.InsertEmail(CreateTestEmail("del@example.com"))

            _repo.PurgeEmailsByIds(New List(Of Integer)() From {id})

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

        <Test>
        Public Sub UpdateHasAttachments_SetsToFalse()
            Dim email As Email = CreateTestEmail("upd-att@example.com")
            email.HasAttachments = True

            Dim id As Integer = _repo.InsertEmail(email)
            Dim before As Email = _repo.GetEmailById(id)
            Assert.IsTrue(before.HasAttachments)

            _repo.UpdateHasAttachments(id, False)
            Dim after As Email = _repo.GetEmailById(id)
            Assert.IsFalse(after.HasAttachments)
        End Sub

        <Test>
        Public Sub UpdateHasAttachments_InBulkMode_SetsToFalse()
            _repo.BeginBulk()
            Dim email As Email = CreateTestEmail("upd-att-bulk@example.com")
            email.HasAttachments = True
            Dim id As Integer = _repo.InsertEmail(email)
            _repo.UpdateHasAttachments(id, False)
            _repo.CommitBulk()

            Dim result As Email = _repo.GetEmailById(id)
            Assert.IsFalse(result.HasAttachments)
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  Prepared Statement 再利用（バルクモード）
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub BulkMode_PreparedStatement_InsertsCorrectly()
            _repo.BeginBulk()
            Dim id1 As Integer = _repo.InsertEmail(CreateTestEmail("ps1@example.com"))
            Dim id2 As Integer = _repo.InsertEmail(CreateTestEmail("ps2@example.com"))
            _repo.CommitBulk()

            Assert.IsTrue(id1 > 0)
            Assert.IsTrue(id2 > id1)
            Assert.AreEqual(2, _repo.GetTotalCount())

            Dim e1 As Email = _repo.GetEmailById(id1)
            Assert.AreEqual("ps1@example.com", e1.MessageId)
            Dim e2 As Email = _repo.GetEmailById(id2)
            Assert.AreEqual("ps2@example.com", e2.MessageId)
        End Sub

        <Test>
        Public Sub BulkMode_PreparedStatement_AttachmentInsertsCorrectly()
            _repo.BeginBulk()
            Dim emailId As Integer = _repo.InsertEmail(CreateTestEmail("psatt@example.com"))
            Dim att As New Attachment()
            att.EmailId = emailId
            att.FileName = "test.pdf"
            att.FilePath = "2025\01\test.pdf"
            att.FileSize = 512
            att.MimeType = "application/pdf"
            Dim attId As Integer = _repo.InsertAttachment(att)
            _repo.CommitBulk()

            Assert.IsTrue(attId > 0)

            ' DB 直接クエリで確認
            Using conn As System.Data.SQLite.SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New System.Data.SQLite.SQLiteCommand(
                    "SELECT file_name FROM attachments WHERE id = @id", conn)
                    cmd.Parameters.AddWithValue("@id", CType(attId, Object))
                    Assert.AreEqual("test.pdf", CStr(cmd.ExecuteScalar()))
                End Using
            End Using
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  最終取り込み日時
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub GetLastImportDate_NoEmails_ReturnsNothing()
            Dim result As DateTime? = _repo.GetLastImportDate()

            Assert.IsFalse(result.HasValue)
        End Sub

        <Test>
        Public Sub GetLastImportDate_WithEmails_ReturnsLatestCreatedAt()
            _repo.InsertEmail(CreateTestEmail("first@example.com"))
            ' created_at は DEFAULT (datetime('now', 'localtime')) で自動設定される
            _repo.InsertEmail(CreateTestEmail("second@example.com"))

            Dim result As DateTime? = _repo.GetLastImportDate()

            Assert.IsTrue(result.HasValue)
            ' created_at は now なので今日の日付であるはず
            Assert.AreEqual(DateTime.Now.Date, result.Value.Date)
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  フォルダ別 MessageID 取得
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub GetMessageIdsByFolder_ReturnsOnlyMatchingFolder()
            Dim email1 As Email = CreateTestEmail("msg1@example.com")
            email1.FolderName = "受信トレイ"
            Dim email2 As Email = CreateTestEmail("msg2@example.com")
            email2.FolderName = "受信トレイ"
            Dim email3 As Email = CreateTestEmail("msg3@example.com")
            email3.FolderName = "送信済み"

            _repo.InsertEmail(email1)
            _repo.InsertEmail(email2)
            _repo.InsertEmail(email3)

            Dim result As Dictionary(Of String, Integer) = _repo.GetMessageIdsByFolder("受信トレイ")

            Assert.AreEqual(2, result.Count)
            Assert.IsTrue(result.ContainsKey("msg1@example.com"))
            Assert.IsTrue(result.ContainsKey("msg2@example.com"))
            Assert.IsFalse(result.ContainsKey("msg3@example.com"))
        End Sub

        <Test>
        Public Sub GetMessageIdsByFolder_EmptyFolder_ReturnsEmptyDictionary()
            Dim result As Dictionary(Of String, Integer) = _repo.GetMessageIdsByFolder("存在しないフォルダ")

            Assert.AreEqual(0, result.Count)
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  一括物理削除（パージ）
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub PurgeEmailsByIds_RemovesEmailsAndRecordsTombstones()
            Dim id1 As Integer = _repo.InsertEmail(CreateTestEmail("del1@example.com"))
            Dim id2 As Integer = _repo.InsertEmail(CreateTestEmail("del2@example.com"))
            Dim id3 As Integer = _repo.InsertEmail(CreateTestEmail("keep@example.com"))

            Dim paths As List(Of String) = _repo.PurgeEmailsByIds(New List(Of Integer)() From {id1, id2})

            ' 削除されていること
            Assert.IsNull(_repo.GetEmailById(id1))
            Assert.IsNull(_repo.GetEmailById(id2))
            ' 残っていること
            Assert.IsNotNull(_repo.GetEmailById(id3))
            ' トゥームストーンに記録されていること
            Assert.IsTrue(_repo.IsMessageIdDeleted("del1@example.com"))
            Assert.IsTrue(_repo.IsMessageIdDeleted("del2@example.com"))
            Assert.IsFalse(_repo.IsMessageIdDeleted("keep@example.com"))
        End Sub

        <Test>
        Public Sub PurgeEmailsByIds_ReturnsAttachmentPaths()
            Dim emailId As Integer = _repo.InsertEmail(CreateTestEmail("att-del@example.com"))
            Dim att As New Attachment()
            att.EmailId = emailId
            att.FileName = "test.pdf"
            att.FilePath = "subdir\test.pdf"
            att.FileSize = 1024
            _repo.InsertAttachment(att)

            Dim paths As List(Of String) = _repo.PurgeEmailsByIds(New List(Of Integer)() From {emailId})

            Assert.AreEqual(1, paths.Count)
            Assert.IsTrue(paths(0).EndsWith("subdir\test.pdf"))
        End Sub

        <Test>
        Public Sub PurgeEmailsByIds_EmptyList_NoError()
            Dim paths As List(Of String) = _repo.PurgeEmailsByIds(New List(Of Integer)())

            Assert.AreEqual(0, paths.Count)
        End Sub

        <Test>
        Public Sub PurgeEmailsByIds_CascadeDeletesAttachments()
            Dim emailId As Integer = _repo.InsertEmail(CreateTestEmail("cascade@example.com"))
            Dim att As New Attachment()
            att.EmailId = emailId
            att.FileName = "doc.txt"
            att.FilePath = "subdir\doc.txt"
            att.FileSize = 512
            _repo.InsertAttachment(att)

            _repo.PurgeEmailsByIds(New List(Of Integer)() From {emailId})

            ' attachments テーブルからも削除されていること
            Using conn As System.Data.SQLite.SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New System.Data.SQLite.SQLiteCommand(
                    "SELECT COUNT(1) FROM attachments WHERE email_id = @id", conn)
                    cmd.Parameters.AddWithValue("@id", CType(emailId, Object))
                    Assert.AreEqual(0, Convert.ToInt32(cmd.ExecuteScalar()))
                End Using
            End Using
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  PurgeExpiredTrash
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub PurgeExpiredTrash_ZeroDays_DoesNothing()
            Dim id As Integer = _repo.InsertEmail(CreateTestEmail("trash@example.com"))
            _repo.SoftDeleteEmail(id)

            Dim paths As List(Of String) = _repo.PurgeExpiredTrash(0)

            Assert.AreEqual(0, paths.Count)
            Assert.AreEqual(1, _repo.GetTrashCount())
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  エラー除外 MessageID
        ' ════════════════════════════════════════════════════════════

        <Test>
        Public Sub GetAllErrorMessageIds_ReturnsRegisteredIds()
            _repo.InsertErrorMessageId("err1@example.com", "受信トレイ", "件名1", "エラー内容1")
            _repo.InsertErrorMessageId("err2@example.com", "送信済み", "件名2", "エラー内容2")

            Dim errorIds As HashSet(Of String) = _repo.GetAllErrorMessageIds()

            Assert.AreEqual(2, errorIds.Count)
            Assert.IsTrue(errorIds.Contains("err1@example.com"))
            Assert.IsTrue(errorIds.Contains("err2@example.com"))
        End Sub

        <Test>
        Public Sub GetAllErrorMessageIds_CaseInsensitive()
            _repo.InsertErrorMessageId("ERR@EXAMPLE.COM", "受信トレイ", "件名", "エラー")

            Dim errorIds As HashSet(Of String) = _repo.GetAllErrorMessageIds()

            Assert.IsTrue(errorIds.Contains("err@example.com"))
        End Sub

        <Test>
        Public Sub InsertErrorMessageId_EmptyId_DoesNotInsert()
            _repo.InsertErrorMessageId("", "受信トレイ", "件名", "エラー")

            Dim errorIds As HashSet(Of String) = _repo.GetAllErrorMessageIds()

            Assert.AreEqual(0, errorIds.Count)
        End Sub

        <Test>
        Public Sub InsertErrorMessageId_DuplicateId_Replaces()
            _repo.InsertErrorMessageId("err@example.com", "受信トレイ", "件名1", "エラー1")
            _repo.InsertErrorMessageId("err@example.com", "受信トレイ", "件名1", "エラー2")

            Dim errorIds As HashSet(Of String) = _repo.GetAllErrorMessageIds()

            Assert.AreEqual(1, errorIds.Count)
        End Sub

        <Test>
        Public Sub InsertErrorMessageId_WithReceivedDate_Persists()
            Dim received As New DateTime(2025, 6, 15, 10, 30, 0)
            _repo.InsertErrorMessageId("err@example.com", "受信トレイ", "件名", "エラー", received, "送信者太郎")

            Dim dt As System.Data.DataTable = _repo.GetErrorMessageEntries()

            Assert.AreEqual(1, dt.Rows.Count)
            Assert.AreEqual("err@example.com", CStr(dt.Rows(0)("message_id")))
            Assert.AreEqual("受信トレイ", CStr(dt.Rows(0)("folder_name")))
            Assert.AreEqual("件名", CStr(dt.Rows(0)("subject")))
            Assert.AreEqual("エラー", CStr(dt.Rows(0)("error_message")))
            Assert.AreEqual("送信者太郎", CStr(dt.Rows(0)("sender_name")))
        End Sub

        <Test>
        Public Sub DeleteErrorMessageId_RemovesSpecificId()
            _repo.InsertErrorMessageId("err1@example.com", "受信トレイ", "件名1", "エラー1")
            _repo.InsertErrorMessageId("err2@example.com", "受信トレイ", "件名2", "エラー2")

            _repo.DeleteErrorMessageId("err1@example.com")

            Dim errorIds As HashSet(Of String) = _repo.GetAllErrorMessageIds()
            Assert.AreEqual(1, errorIds.Count)
            Assert.IsFalse(errorIds.Contains("err1@example.com"))
            Assert.IsTrue(errorIds.Contains("err2@example.com"))
        End Sub

        <Test>
        Public Sub ClearAllErrorMessageIds_RemovesAll()
            _repo.InsertErrorMessageId("err1@example.com", "受信トレイ", "件名1", "エラー1")
            _repo.InsertErrorMessageId("err2@example.com", "受信トレイ", "件名2", "エラー2")

            _repo.ClearAllErrorMessageIds()

            Dim errorIds As HashSet(Of String) = _repo.GetAllErrorMessageIds()
            Assert.AreEqual(0, errorIds.Count)
        End Sub

        <Test>
        Public Sub GetErrorMessageIdCount_ReturnsCorrectCount()
            Assert.AreEqual(0, _repo.GetErrorMessageIdCount())

            _repo.InsertErrorMessageId("err1@example.com", "受信トレイ", "件名1", "エラー1")
            Assert.AreEqual(1, _repo.GetErrorMessageIdCount())

            _repo.InsertErrorMessageId("err2@example.com", "受信トレイ", "件名2", "エラー2")
            Assert.AreEqual(2, _repo.GetErrorMessageIdCount())
        End Sub

        <Test>
        Public Sub InsertErrorMessageId_InBulkMode_Persists()
            _repo.BeginBulk()
            _repo.InsertErrorMessageId("bulk-err@example.com", "受信トレイ", "件名", "エラー")
            _repo.CommitBulk()

            Dim errorIds As HashSet(Of String) = _repo.GetAllErrorMessageIds()
            Assert.AreEqual(1, errorIds.Count)
            Assert.IsTrue(errorIds.Contains("bulk-err@example.com"))
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  ヘルパー
        ' ════════════════════════════════════════════════════════════

        ' ── 同一 MessageID のフォルダ別保存 ──────────────────────────

        <Test>
        Public Sub InsertEmail_SameMessageIdDifferentFolder_BothInserted()
            Dim email1 As Email = CreateTestEmail("shared@example.com")
            email1.FolderName = "受信トレイ"
            Dim email2 As Email = CreateTestEmail("shared@example.com")
            email2.FolderName = "送信済みアイテム"

            Dim id1 As Integer = _repo.InsertEmail(email1)
            Dim id2 As Integer = _repo.InsertEmail(email2)

            Assert.IsTrue(id1 > 0)
            Assert.IsTrue(id2 > 0)
            Assert.AreNotEqual(id1, id2)
        End Sub

        <Test>
        Public Sub GetAllMessageIdFolderPairs_ReturnsPairsPerFolder()
            Dim email1 As Email = CreateTestEmail("pair@example.com")
            email1.FolderName = "受信トレイ"
            Dim email2 As Email = CreateTestEmail("pair@example.com")
            email2.FolderName = "送信済みアイテム"

            _repo.InsertEmail(email1)
            _repo.InsertEmail(email2)

            Dim pairs As HashSet(Of String) = _repo.GetAllMessageIdFolderPairs()

            Assert.IsTrue(pairs.Contains("pair@example.com" & vbTab & "受信トレイ"))
            Assert.IsTrue(pairs.Contains("pair@example.com" & vbTab & "送信済みアイテム"))
            Assert.That(pairs.Count, [Is].GreaterThanOrEqualTo(2))
        End Sub

        <Test>
        Public Sub InsertEmail_SameMessageIdSameFolder_ThrowsException()
            Dim email1 As Email = CreateTestEmail("dup-folder@example.com")
            email1.FolderName = "受信トレイ"
            _repo.InsertEmail(email1)

            Dim email2 As Email = CreateTestEmail("dup-folder@example.com")
            email2.FolderName = "受信トレイ"

            Assert.Throws(Of System.Data.SQLite.SQLiteException)(
                Sub() _repo.InsertEmail(email2))
        End Sub

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
