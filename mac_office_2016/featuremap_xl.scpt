FasdUAS 1.101.10   ��   ��    k             l     ����  r       	  n     
  
 I    �� ���� "0 readfeaturefile readFeatureFile   ��  m       �   � M a c i n t o s h   H D : U s e r s : m a c : D o c u m e n t s : s o u r c e : a p p l e s c r i p t : o m n i g r a f f l e : t e x t 2 p r o c e s s : f e a t u r e s : 1 - d r a w - p r o c e s s - - - c r e a t e - l a n e . f e a t u r e��  ��     f      	 o      ���� 0 vdummy vDummy��  ��        l     ��������  ��  ��        l     ��  ��    ] W---------------------------------------------------------------------------------------     �   � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -      l     ��  ��    ? 9 description: ask user where to expect the .feature files     �   r   d e s c r i p t i o n :   a s k   u s e r   w h e r e   t o   e x p e c t   t h e   . f e a t u r e   f i l e s      l     ��  ��    _ Y parameters:		pDummy		- it seems that Excel 2016 expectall funtions to have one parameter     �     �   p a r a m e t e r s : 	 	 p D u m m y 	 	 -   i t   s e e m s   t h a t   E x c e l   2 0 1 6   e x p e c t a l l   f u n t i o n s   t o   h a v e   o n e   p a r a m e t e r   ! " ! l     �� # $��   # ' ! return value: has to be a string    $ � % % B   r e t u r n   v a l u e :   h a s   t o   b e   a   s t r i n g "  & ' & l     �� ( )��   ( ] W---------------------------------------------------------------------------------------    ) � * * � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - '  + , + i      - . - I      �� /���� *0 choosefeaturefolder chooseFeatureFolder /  0�� 0 o      ���� 0 pdummy pDummy��  ��   . Q     8 1 2 3 1 O    , 4 5 4 k    + 6 6  7 8 7 I   �� 9��
�� .miscactvnull��� ��� obj  9 m    ��
�� 
capp��   8  : ; : r     < = < l    >���� > I   ���� ?
�� .sysostflalis    ��� null��   ? �� @ A
�� 
prmp @ m     B B � C C * c h o o s e   f e a t u r e   f o l d e r A �� D��
�� 
dflc D l    E���� E I   �� F G
�� .earsffdralis        afdr F l    H���� H m    ��
�� afdmdesk��  ��   G �� I��
�� 
from I m    ��
�� fldmfldu��  ��  ��  ��  ��  ��   = o      ���� 0 vpath vPath ;  J�� J L    + K K b    * L M L b    $ N O N n    " P Q P 1     "��
�� 
pURL Q o     ���� 0 vpath vPath O m   " # R R � S S  # @ # @ M n   $ ) T U T 1   ' )��
�� 
dnam U n   $ ' V W V m   % '��
�� 
cdis W o   $ %���� 0 vpath vPath��   5 m     X X�                                                                                  MACS  alis    t  Macintosh HD               ѿF�H+   (B�
Finder.app                                                      *����~        ����  	                CoreServices    ѿ*n      ��o�     (B� (B� (B�  6Macintosh HD:System: Library: CoreServices: Finder.app   
 F i n d e r . a p p    M a c i n t o s h   H D  &System/Library/CoreServices/Finder.app  / ��   2 R      ������
�� .ascrerr ****      � ****��  ��   3 L   4 8 Y Y m   4 7 Z Z � [ [   ,  \ ] \ l     ��������  ��  ��   ]  ^ _ ^ l     �� ` a��   ` ] W---------------------------------------------------------------------------------------    a � b b � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - _  c d c l     �� e f��   e ; 5 description: read file names from the feature folder    f � g g j   d e s c r i p t i o n :   r e a d   f i l e   n a m e s   f r o m   t h e   f e a t u r e   f o l d e r d  h i h l     �� j k��   j U O parameters:		pFeatureFolderPath		- the directory containing all .feature files    k � l l �   p a r a m e t e r s : 	 	 p F e a t u r e F o l d e r P a t h 	 	 -   t h e   d i r e c t o r y   c o n t a i n i n g   a l l   . f e a t u r e   f i l e s i  m n m l     �� o p��   o 6 0 return value: the .feature file names as string    p � q q `   r e t u r n   v a l u e :   t h e   . f e a t u r e   f i l e   n a m e s   a s   s t r i n g n  r s r l     �� t u��   t ] W---------------------------------------------------------------------------------------    u � v v � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - s  w x w i     y z y I      �� {���� *0 getfeaturefilenames getFeatureFileNames {  |�� | o      ���� (0 pfeaturefolderpath pFeatureFolderPath��  ��   z k     G } }  ~  ~ r      � � � J     ����   � o      ���� &0 vfeaturefilenames vFeatureFileNames   � � � O    < � � � k   	 ; � �  � � � r   	  � � � c   	  � � � o   	 
���� (0 pfeaturefolderpath pFeatureFolderPath � m   
 ��
�� 
alis � o      ���� "0 vfeaturesfolder vFeaturesFolder �  � � � r     � � � l    ����� � e     � � 6    � � � n     � � � 2   ��
�� 
file � o    ���� "0 vfeaturesfolder vFeaturesFolder � D     � � � 1    ��
�� 
pnam � m     � � � � �  . f e a t u r e��  ��   � o      ���� 0 vfeaturefiles vFeatureFiles �  ��� � X    ; ��� � � r   / 6 � � � e   / 3 � � n   / 3 � � � 1   0 2��
�� 
pURL � o   / 0���� 0 vfeaturefile vFeatureFile � n       � � �  ;   4 5 � o   3 4���� &0 vfeaturefilenames vFeatureFileNames�� 0 vfeaturefile vFeatureFile � o   " #���� 0 vfeaturefiles vFeatureFiles��   � m     � ��                                                                                  MACS  alis    t  Macintosh HD               ѿF�H+   (B�
Finder.app                                                      *����~        ����  	                CoreServices    ѿ*n      ��o�     (B� (B� (B�  6Macintosh HD:System: Library: CoreServices: Finder.app   
 F i n d e r . a p p    M a c i n t o s h   H D  &System/Library/CoreServices/Finder.app  / ��   �  � � � r   = B � � � m   = > � � � � �  # @ # @ � n      � � � 1   ? A��
�� 
txdl � 1   > ?��
�� 
ascr �  ��� � L   C G � � c   C F � � � o   C D���� &0 vfeaturefilenames vFeatureFileNames � m   D E��
�� 
TEXT��   x  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � ] W---------------------------------------------------------------------------------------    � � � � � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - �  � � � l     �� � ���   � F @ description: read the content from a given single .feature file    � � � � �   d e s c r i p t i o n :   r e a d   t h e   c o n t e n t   f r o m   a   g i v e n   s i n g l e   . f e a t u r e   f i l e �  � � � l     �� � ���   � U O parameters:		pFeatureFolderPath		- the directory containing all .feature files    � � � � �   p a r a m e t e r s : 	 	 p F e a t u r e F o l d e r P a t h 	 	 -   t h e   d i r e c t o r y   c o n t a i n i n g   a l l   . f e a t u r e   f i l e s �  � � � l     �� � ���   � 6 0 return value: the .feature file names as string    � � � � `   r e t u r n   v a l u e :   t h e   . f e a t u r e   f i l e   n a m e s   a s   s t r i n g �  � � � l     �� � ���   � ] W---------------------------------------------------------------------------------------    � � � � � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - �  � � � i     � � � I      �� ����� "0 readfeaturefile readFeatureFile �  ��� � o      ���� $0 pfeaturefilepath pFeatureFilePath��  ��   � k     $ � �  � � � l     ��������  ��  ��   �  � � � q       � � ������ (0 voldtextdelimiters vOldTextDelimiters��   �  � � � q       � � ������ 0 vfeaturetext vFeatureText��   �  � � � l     ��������  ��  ��   �  � � � r      � � � n     � � � 1    ��
�� 
txdl � 1     ��
�� 
ascr � o      ���� (0 voldtextdelimiters vOldTextDelimiters �  � � � r     � � � m     � � � � �  # @ # @ � n      � � � 1    
��
�� 
txdl � 1    ��
�� 
ascr �  � � � r     � � � c     � � � l    ����� � n     � � � 2   ��
�� 
cpar � l    ����� � I   �� � �
�� .rdwrread****        **** � l    ����� � c     � � � o    ���� $0 pfeaturefilepath pFeatureFilePath � m    ��
�� 
alis��  ��   � �� ���
�� 
as   � m    ��
�� 
utf8��  ��  ��  ��  ��   � m    ��
�� 
TEXT � o      ���� 0 vfeaturetext vFeatureText �  � � � r    ! � � � o    ���� (0 voldtextdelimiters vOldTextDelimiters � n        1     ��
�� 
txdl 1    ��
�� 
ascr � �� L   " $ o   " #�� 0 vfeaturetext vFeatureText��   � �~ l     �}�|�{�}  �|  �{  �~       �z	�z   �y�x�w�v�y *0 choosefeaturefolder chooseFeatureFolder�x *0 getfeaturefilenames getFeatureFileNames�w "0 readfeaturefile readFeatureFile
�v .aevtoappnull  �   � **** �u .�t�s
�r�u *0 choosefeaturefolder chooseFeatureFolder�t �q�q   �p�p 0 pdummy pDummy�s  
 �o�n�o 0 pdummy pDummy�n 0 vpath vPath  X�m�l�k B�j�i�h�g�f�e�d�c R�b�a�`�_ Z
�m 
capp
�l .miscactvnull��� ��� obj 
�k 
prmp
�j 
dflc
�i afdmdesk
�h 
from
�g fldmfldu
�f .earsffdralis        afdr�e 
�d .sysostflalis    ��� null
�c 
pURL
�b 
cdis
�a 
dnam�`  �_  �r 9 .� &�j O*������l 	� E�O��,�%��,�,%UW X  a  �^ z�]�\�[�^ *0 getfeaturefilenames getFeatureFileNames�] �Z�Z   �Y�Y (0 pfeaturefolderpath pFeatureFolderPath�\   �X�W�V�U�T�X (0 pfeaturefolderpath pFeatureFolderPath�W &0 vfeaturefilenames vFeatureFileNames�V "0 vfeaturesfolder vFeaturesFolder�U 0 vfeaturefiles vFeatureFiles�T 0 vfeaturefile vFeatureFile  ��S�R�Q ��P�O�N�M ��L�K�J
�S 
alis
�R 
file  
�Q 
pnam
�P 
kocl
�O 
cobj
�N .corecnte****       ****
�M 
pURL
�L 
ascr
�K 
txdl
�J 
TEXT�[ HjvE�O� 4��&E�O��-�[�,\Z�?1EE�O �[��l kh ��,E�6F[OY��UO���,FO��& �I ��H�G�F�I "0 readfeaturefile readFeatureFile�H �E�E   �D�D $0 pfeaturefilepath pFeatureFilePath�G   �C�B�A�C $0 pfeaturefilepath pFeatureFilePath�B (0 voldtextdelimiters vOldTextDelimiters�A 0 vfeaturetext vFeatureText 	�@�? ��>�=�<�;�:�9
�@ 
ascr
�? 
txdl
�> 
alis
�= 
as  
�< 
utf8
�; .rdwrread****        ****
�: 
cpar
�9 
TEXT�F %��,E�O���,FO��&��l �-�&E�O���,FO�	 �8�7�6�5
�8 .aevtoappnull  �   � **** k       �4�4  �7  �6      �3�2�3 "0 readfeaturefile readFeatureFile�2 0 vdummy vDummy�5 	)�k+ E� ascr  ��ޭ