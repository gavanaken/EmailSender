   M S B A M L   `   `   `  ����M  IEmailSender, Version=1.0.7358.28932, Culture=neutral, PublicKeyToken=null)clr-namespace:EmailSenderEmailSender  5   	   ^ ZSpiceLogicWPFHtmlEditor, Version=2.7.5.0, Culture=neutral, PublicKeyToken=8cd3285c24067f7fdGclr-namespace:SpiceLogic.WpfHtmlEditor;assembly=SpiceLogicWPFHtmlEditorSpiceLogic.WpfHtmlEditor 5   	       EmailSender.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   +localclr-namespace:EmailSender     5   	   aWpfHtmlEditorGclr-namespace:SpiceLogic.WpfHtmlEditor;assembly=SpiceLogicWPFHtmlEditor     5   	     � Title$  Email Sender��5	   	   $	��772��6   $
��1276��6+    � ShowInTaskbar .   � 
ResizeMode$ CanMinimize=�6N   ���	#FF1B1A1A
�������.��5
      � $��0,0,-6,0q�6    � ColumnDefinitions 5   
   �� 5      $
��899* �6   �� 5      $
��489* �6   �� 5      $��0* �6   .+�5   
   ,  &SpiceLogic.WpfHtmlEditor.WpfHtmlEditor  -   5          Name$ Html_Editor��5   &   $
��Left=�6;   $	��Top=�6V   $��	397,0,0,0q�6n   $	��866��6�   $��2��6�   $	��713��6�   &   System.Windows.Controls.DataGrid  5   
   -   5          Name$ 
Group_List��5      $
��Left=�6'   $	��496��6B   $��	2,217,0,0q�6O   $	��Top=�6b   $	��390��6z    � 	Resources 5      )  #System.Windows.Controls.DataGridRow �� 
TargetType��%E   '        �� 5      "  6   .�5      -   5        5      5      5      G� 5   
   -   5         G�Name$ Password_Box��5      $
��Left=�6,   $��
158,97,0,0q�6G   $	��Top=�6[   $	��189��6s   $��22��6   �� 5   
   -   5        	 ��Name$	 	Email_Box��5      $
��Left=�6%   $��23��6@   $��
158,69,0,0q�6L   
 �� TextWrapping$

 Wrap=�6`   $	��Top=�6t   $	��189��6�   6�   �� 5   
   $��Email Address:��6   $
��Left=�6)   $��	40,66,0,0q�6D   $	��Top=�6W   
�������$��92��6�   $��26��6�   �� 5   
   $��	Password:��6   $
��Left=�6$   $��	40,97,0,0q�6?   $	��Top=�6R   
�������$��92��6}   $��26��6�   �� 5   
   -   5        $	 Substitution_String��5      $
��Left=�6/   $��158,126,0,0q�6J   $	��Top=�6_   $	��189��6w   $��22��6�   �� 5   
   $��Substitution String:��6   $
��Left=�6/   $��
40,124,0,0q�6J   $	��Top=�6^   
�������$	��123��6�   $��26��6�   �� 5   
   $��
Group List��6   $
��Left=�6%   $��	2,186,0,0q�6@   $	��Top=�6S   
�������$��98��6~   $��32��6�   $��16(�6�   $
��Bold"�6�   �� 5    
   -   5         ��Name$ From_CSV��5       $��From CSV��6#   $
��Left=�66   $��158,192,0,0q�6Q   $	��Top=�6f   $��76��6~   $��20��6�   �� 5!   
   -   5        $ 	Send_Test��5!      $��	Send Test��6$   $
��Left=�68   $��
272,25,0,0q�6S   $	��Top=�6g   $��75��6   $��20��6�   �� 5"   
   -	   5        $ Send_Emails��5"      $��Send Emails!��6&   $
��Left=�6=   $��
158,25,0,0q�6X   $	��Top=�6l   $��75��6�   $��20��6�   ]� 5#   
   $
��Left=�6   $��30��6/   ��.  $	��Top=�6Q   $
��1271��6i   $��	0,713,0,0q�6v   $��2��6�   ���	#FF211E1E.+�5$      �� -
   5         ��Name$ 	StatusBar��5$      $��30��6+   $	��Top=�67   $	��678��6O   ���	#FF1B1A1A
�������.@�5%      �� -   5         ��Name$ StatusBarMessage��5%      $
��Wrap=�66   $�� ��6J   $��24��6R   $	��365��6^   � 5&      -   5         �Name$ Progress_Bar��5&      $��15��64   $	��Top=�6@   $	��288��6X   $y�Hidden=�6d    J� RenderTransformOrigin$ 0.453,-0.35&�6x   5'      �� 5(      $��30��6   $��NoWrap=�6$   $l��f            Copyright © 2020. Greg Van Aken. Made For the Haverford College 'Ford S-Chords         		��6:   $	��Top=�6�   $	��510��6�   $��Right=�6�   �� 5)      -   5        $ Help_Button��5)      $
��Help��6*   $
��Left=�69   $	��Top=�6T   $��67��6l   $ 	1.561,3.3&�6w   $��Right=�6�   6�   5*      �� 5+   
   $
��Left=�6   $��67��6+   $	��Top=�67   $	��132��6O   $��/images/Logo_gray.png��6[   $��	16,-1,0,0q�6z   �� 5,   
   -   5        $	 Subject_Box��5,      $
��Left=�6'   $��158,157,0,0q�6B   $	��Top=�6W   $	��189��6o   $��22��6{   �� 5-   
   $��Subject:��6   $
��Left=�6#   $��
40,155,0,0q�6>   $	��Top=�6R   
�������$	��123��6}   $��26��6�   �� 5.   
   -   5        $ Preview��5.      $��Preview��6"   $
��Left=�64   $��271,192,0,0q�6O   $	��Top=�6d   $��76��6|   $��20��6�   5/      50      