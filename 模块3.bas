Attribute VB_Name = "ģ��3"
Sub ����ȫ��ͼƬ��С() '����ͼƬ�ߴ�

Title = "����ȫ��ͼƬ��С������¿�һ���ĵ��ٴ���"

Message = "����ͼƬ��ȣ���λ����"

MyValue = InputBox(Message, Title, Default)

MyValue = MyValue * 28.35

Dim n 'ͼƬ����

On Error Resume Next '���Դ���

For n = 1 To ActiveDocument.InlineShapes.Count 'InlineShapes ���� ͼƬ

ActiveDocument.InlineShapes(n).Width = MyValue '����ͼƬ��� 10cm�����У�Word��1cm=28.35px

Next n
    MsgBox "�ߴ��������"
End Sub
