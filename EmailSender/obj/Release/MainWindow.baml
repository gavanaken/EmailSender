   M S B A M L   `   `   `  ���� M  IEmailSender, Version=1.0.7358.29220, Culture=neutral, PublicKeyToken=null)clr-namespace:EmailSenderEmailSender  ^ ZSpiceLogicWPFHtmlEditor, Version=2.7.5.0, Culture=neutral, PublicKeyToken=8cd3285c24067f7fdGclr-namespace:SpiceLogic.WpfHtmlEditor;assembly=SpiceLogicWPFHtmlEditorSpiceLogic.WpfHtmlEditor     EmailSender.MainWindow   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     +localclr-namespace:EmailSender     aWpfHtmlEditorGclr-namespace:SpiceLogic.WpfHtmlEditor;assembly=SpiceLogicWPFHtmlEditor       � Title$  Email Sender��$	��772��$
��1276�� � 
ResizeMode$ CanMinimize=����	#FF1B1A1A
������� � Icon$ Icon.ico��.��� $��0,0,-6,0q� � ColumnDefinitions �� $
��899* ��� $
��489* ��� $��0* �.+�,  &SpiceLogic.WpfHtmlEditor.WpfHtmlEditor  -     Name$ Html_Editor��$
��Left=�$	��Top=�$��	397,0,0,0q�$	��866��$��2��$	��713��&   System.Windows.Controls.DataGrid  -     Name$ 
Group_List��$
��Left=�$	��496��$��	2,217,0,0q�$	��Top=�$	��390�� � 	Resources )  #System.Windows.Controls.DataGridRow �� 
TargetType��%   '        �� "  .�-   G� -    G�Name$ Password_Box��$
��Left=�$��
158,97,0,0q�$	��Top=�$	��189��$��22���� -   	 ��Name$	 	Email_Box��$
��Left=�$��23��$��
158,69,0,0q�
 �� TextWrapping$

 Wrap=�$	��Top=�$	��189���� $��Email Address:��$
��Left=�$��	40,66,0,0q�$	��Top=�
�������$��92��$��26���� $��	Password:��$
��Left=�$��	40,97,0,0q�$	��Top=�
�������$��92��$��26���� -   $	 Substitution_String��$
��Left=�$��158,126,0,0q�$	��Top=�$	��189��$��22���� $��Substitution String:��$
��Left=�$��
40,124,0,0q�$	��Top=�
�������$	��123��$��26���� $��
Group List��$
��Left=�$��	2,186,0,0q�$	��Top=�
�������$��98��$��32��$��16(�$
��Bold"��� -    ��Name$ From_CSV��$��From CSV��$
��Left=�$��158,192,0,0q�$	��Top=�$��76��$��20���� -   $ 	Send_Test��$��	Send Test��$
��Left=�$��
272,25,0,0q�$	��Top=�$��75��$��20���� -	   $ Send_Emails��$��Send Emails!��$
��Left=�$��
158,25,0,0q�$	��Top=�$��75��$��20��]� $
��Left=�$��30����.  $	��Top=�$
��1271��$��	0,713,0,0q�$��2�����	#FF211E1E.+��� -
    ��Name$ 	StatusBar��$��30��$	��Top=�$	��678�����	#FF1B1A1A
�������.@��� -    ��Name$ StatusBarMessage��$
��Wrap=�$�� ��$��24��$	��365��� -    �Name$ Progress_Bar��$��15��$	��Top=�$	��288��$y�Hidden=� J� RenderTransformOrigin$ 0.453,-0.35&��� $��30��$��NoWrap=�$l��f            Copyright © 2020. Greg Van Aken. Made For the Haverford College 'Ford S-Chords         		��$	��Top=�$	��510��$��Right=��� -   $ Help_Button��$
��Help��$
��Left=�$	��Top=�$��67��$ 	1.561,3.3&�$��Right=��� $
��Left=�$��67��$	��Top=�$	��132��$��/images/Logo_gray.png��$��	16,-1,0,0q��� -   $	 Subject_Box��$
��Left=�$��158,157,0,0q�$	��Top=�$	��189��$��22���� $��Subject:��$
��Left=�$��
40,155,0,0q�$	��Top=�
�������$	��123��$��26���� -   $ Preview��$��Preview��$
��Left=�$��271,192,0,0q�$	��Top=�$��76��$��20��