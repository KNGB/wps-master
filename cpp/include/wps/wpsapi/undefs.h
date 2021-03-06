/*
** Copyright @ 2012-2019, Kingsoft office,All rights reserved.
**
** Redistribution and use in source and binary forms ,without modification and
** selling solely, are permitted provided that the following conditions are
** met:
**
** 1.Redistributions of source code must retain the above copyright notice,
**   this list of conditions and the following disclaimer.
** 2.Redistributions in binary form must reproduce the above copyright notice,
**	 this list of conditions and the following disclaimer in the documentation
**	 and/or other materials provided with the distribution.
** 3.Neither the name of the copyright holder nor the names of its contributors
**	 may be used to endorse or promote products derived from this software
**	 without specific prior written permission.
**
** SPECIAL NOTE:THIS SOFTWARE IS NOT PERMITTED TO BE MODIFIED OR SOLD SOLELY AT
** ANY TIME AND UNDER ANY CIRCUMSTANCES, EXCEPT WITH THE WRITTEN PERMISSION OF
** KINGSOFT OFFICE
**
** THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
** AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
** IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
** ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
** LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
** CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
** SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
** INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
** CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
** ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
** POSSIBILITY OF SUCH DAMAGE.
**/
#undef ___Application_FWD_DEFINED__
#undef ___Global_FWD_DEFINED__
#undef __FontNames_FWD_DEFINED__
#undef __Languages_FWD_DEFINED__
#undef __Language_FWD_DEFINED__
#undef __Documents_FWD_DEFINED__
#undef ___Document_FWD_DEFINED__
#undef __Template_FWD_DEFINED__
#undef __Templates_FWD_DEFINED__
#undef __RoutingSlip_FWD_DEFINED__
#undef __Bookmark_FWD_DEFINED__
#undef __Bookmarks_FWD_DEFINED__
#undef __Variable_FWD_DEFINED__
#undef __Variables_FWD_DEFINED__
#undef __RecentFile_FWD_DEFINED__
#undef __RecentFiles_FWD_DEFINED__
#undef __Window_FWD_DEFINED__
#undef __Windows_FWD_DEFINED__
#undef __Pane_FWD_DEFINED__
#undef __Panes_FWD_DEFINED__
#undef __Range_FWD_DEFINED__
#undef __ListFormat_FWD_DEFINED__
#undef __Find_FWD_DEFINED__
#undef __Replacement_FWD_DEFINED__
#undef __Characters_FWD_DEFINED__
#undef __Words_FWD_DEFINED__
#undef __Sentences_FWD_DEFINED__
#undef __Sections_FWD_DEFINED__
#undef __Section_FWD_DEFINED__
#undef __Paragraphs_FWD_DEFINED__
#undef __Paragraph_FWD_DEFINED__
#undef __DropCap_FWD_DEFINED__
#undef __TabStops_FWD_DEFINED__
#undef __TabStop_FWD_DEFINED__
#undef ___ParagraphFormat_FWD_DEFINED__
#undef ___Font_FWD_DEFINED__
#undef __Table_FWD_DEFINED__
#undef __Row_FWD_DEFINED__
#undef __Column_FWD_DEFINED__
#undef __Cell_FWD_DEFINED__
#undef __Tables_FWD_DEFINED__
#undef __Rows_FWD_DEFINED__
#undef __Columns_FWD_DEFINED__
#undef __Cells_FWD_DEFINED__
#undef __AutoCorrect_FWD_DEFINED__
#undef __AutoCorrectEntries_FWD_DEFINED__
#undef __AutoCorrectEntry_FWD_DEFINED__
#undef __FirstLetterExceptions_FWD_DEFINED__
#undef __FirstLetterException_FWD_DEFINED__
#undef __TwoInitialCapsExceptions_FWD_DEFINED__
#undef __TwoInitialCapsException_FWD_DEFINED__
#undef __Footnotes_FWD_DEFINED__
#undef __Endnotes_FWD_DEFINED__
#undef __Comments_FWD_DEFINED__
#undef __Footnote_FWD_DEFINED__
#undef __Endnote_FWD_DEFINED__
#undef __Comment_FWD_DEFINED__
#undef __Borders_FWD_DEFINED__
#undef __Border_FWD_DEFINED__
#undef __Shading_FWD_DEFINED__
#undef __TextRetrievalMode_FWD_DEFINED__
#undef __AutoTextEntries_FWD_DEFINED__
#undef __AutoTextEntry_FWD_DEFINED__
#undef __System_FWD_DEFINED__
#undef __OLEFormat_FWD_DEFINED__
#undef __LinkFormat_FWD_DEFINED__
#undef ___OLEControl_FWD_DEFINED__
#undef __Fields_FWD_DEFINED__
#undef __Field_FWD_DEFINED__
#undef __Browser_FWD_DEFINED__
#undef __Styles_FWD_DEFINED__
#undef __Style_FWD_DEFINED__
#undef __Frames_FWD_DEFINED__
#undef __Frame_FWD_DEFINED__
#undef __FormFields_FWD_DEFINED__
#undef __FormField_FWD_DEFINED__
#undef __TextInput_FWD_DEFINED__
#undef __CheckBox_FWD_DEFINED__
#undef __DropDown_FWD_DEFINED__
#undef __ListEntries_FWD_DEFINED__
#undef __ListEntry_FWD_DEFINED__
#undef __TablesOfFigures_FWD_DEFINED__
#undef __TableOfFigures_FWD_DEFINED__
#undef __MailMerge_FWD_DEFINED__
#undef __MailMergeFields_FWD_DEFINED__
#undef __MailMergeField_FWD_DEFINED__
#undef __MailMergeDataSource_FWD_DEFINED__
#undef __MailMergeFieldNames_FWD_DEFINED__
#undef __MailMergeFieldName_FWD_DEFINED__
#undef __MailMergeDataFields_FWD_DEFINED__
#undef __MailMergeDataField_FWD_DEFINED__
#undef __Envelope_FWD_DEFINED__
#undef __MailingLabel_FWD_DEFINED__
#undef __CustomLabels_FWD_DEFINED__
#undef __CustomLabel_FWD_DEFINED__
#undef __TablesOfContents_FWD_DEFINED__
#undef __TableOfContents_FWD_DEFINED__
#undef __TablesOfAuthorities_FWD_DEFINED__
#undef __TableOfAuthorities_FWD_DEFINED__
#undef __Dialogs_FWD_DEFINED__
#undef __Dialog_FWD_DEFINED__
#undef __PageSetup_FWD_DEFINED__
#undef __LineNumbering_FWD_DEFINED__
#undef __TextColumns_FWD_DEFINED__
#undef __TextColumn_FWD_DEFINED__
#undef __Selection_FWD_DEFINED__
#undef __TablesOfAuthoritiesCategories_FWD_DEFINED__
#undef __TableOfAuthoritiesCategory_FWD_DEFINED__
#undef __CaptionLabels_FWD_DEFINED__
#undef __CaptionLabel_FWD_DEFINED__
#undef __AutoCaptions_FWD_DEFINED__
#undef __AutoCaption_FWD_DEFINED__
#undef __Indexes_FWD_DEFINED__
#undef __Index_FWD_DEFINED__
#undef __AddIn_FWD_DEFINED__
#undef __AddIns_FWD_DEFINED__
#undef __Revisions_FWD_DEFINED__
#undef __Revision_FWD_DEFINED__
#undef __Task_FWD_DEFINED__
#undef __Tasks_FWD_DEFINED__
#undef __HeadersFooters_FWD_DEFINED__
#undef __HeaderFooter_FWD_DEFINED__
#undef __PageNumbers_FWD_DEFINED__
#undef __PageNumber_FWD_DEFINED__
#undef __Subdocuments_FWD_DEFINED__
#undef __Subdocument_FWD_DEFINED__
#undef __HeadingStyles_FWD_DEFINED__
#undef __HeadingStyle_FWD_DEFINED__
#undef __StoryRanges_FWD_DEFINED__
#undef __ListLevel_FWD_DEFINED__
#undef __ListLevels_FWD_DEFINED__
#undef __ListTemplate_FWD_DEFINED__
#undef __ListTemplates_FWD_DEFINED__
#undef __ListParagraphs_FWD_DEFINED__
#undef __List_FWD_DEFINED__
#undef __Lists_FWD_DEFINED__
#undef __ListGallery_FWD_DEFINED__
#undef __ListGalleries_FWD_DEFINED__
#undef __KeyBindings_FWD_DEFINED__
#undef __KeysBoundTo_FWD_DEFINED__
#undef __KeyBinding_FWD_DEFINED__
#undef __FileConverter_FWD_DEFINED__
#undef __FileConverters_FWD_DEFINED__
#undef __SynonymInfo_FWD_DEFINED__
#undef __Hyperlinks_FWD_DEFINED__
#undef __Hyperlink_FWD_DEFINED__
#undef __Shapes_FWD_DEFINED__
#undef __ShapeRange_FWD_DEFINED__
#undef __GroupShapes_FWD_DEFINED__
#undef __Shape_FWD_DEFINED__
#undef __TextFrame_FWD_DEFINED__
#undef ___LetterContent_FWD_DEFINED__
#undef __View_FWD_DEFINED__
#undef __Zoom_FWD_DEFINED__
#undef __Zooms_FWD_DEFINED__
#undef __InlineShape_FWD_DEFINED__
#undef __InlineShapes_FWD_DEFINED__
#undef __SpellingSuggestions_FWD_DEFINED__
#undef __SpellingSuggestion_FWD_DEFINED__
#undef __Dictionaries_FWD_DEFINED__
#undef __HangulHanjaConversionDictionaries_FWD_DEFINED__
#undef __Dictionary_FWD_DEFINED__
#undef __ReadabilityStatistics_FWD_DEFINED__
#undef __ReadabilityStatistic_FWD_DEFINED__
#undef __Versions_FWD_DEFINED__
#undef __Version_FWD_DEFINED__
#undef __Options_FWD_DEFINED__
#undef __MailMessage_FWD_DEFINED__
#undef __ProofreadingErrors_FWD_DEFINED__
#undef __Mailer_FWD_DEFINED__
#undef __WrapFormat_FWD_DEFINED__
#undef __HangulAndAlphabetExceptions_FWD_DEFINED__
#undef __HangulAndAlphabetException_FWD_DEFINED__
#undef __Adjustments_FWD_DEFINED__
#undef __CalloutFormat_FWD_DEFINED__
#undef __ColorFormat_FWD_DEFINED__
#undef __ConnectorFormat_FWD_DEFINED__
#undef __FillFormat_FWD_DEFINED__
#undef __FreeformBuilder_FWD_DEFINED__
#undef __LineFormat_FWD_DEFINED__
#undef __PictureFormat_FWD_DEFINED__
#undef __ShadowFormat_FWD_DEFINED__
#undef __ShapeNode_FWD_DEFINED__
#undef __ShapeNodes_FWD_DEFINED__
#undef __TextEffectFormat_FWD_DEFINED__
#undef __ThreeDFormat_FWD_DEFINED__
#undef __ApplicationEvents_FWD_DEFINED__
#undef __DocumentEvents_FWD_DEFINED__
#undef __OCXEvents_FWD_DEFINED__
#undef __IApplicationEvents_FWD_DEFINED__
#undef __IApplicationEvents2_FWD_DEFINED__
#undef __ApplicationEvents2_FWD_DEFINED__
#undef __EmailAuthor_FWD_DEFINED__
#undef __EmailOptions_FWD_DEFINED__
#undef __EmailSignature_FWD_DEFINED__
#undef __Email_FWD_DEFINED__
#undef __HorizontalLineFormat_FWD_DEFINED__
#undef __Frameset_FWD_DEFINED__
#undef __DefaultWebOptions_FWD_DEFINED__
#undef __WebOptions_FWD_DEFINED__
#undef __OtherCorrectionsExceptions_FWD_DEFINED__
#undef __OtherCorrectionsException_FWD_DEFINED__
#undef __EmailSignatureEntries_FWD_DEFINED__
#undef __EmailSignatureEntry_FWD_DEFINED__
#undef __HTMLDivision_FWD_DEFINED__
#undef __HTMLDivisions_FWD_DEFINED__
#undef __DiagramNode_FWD_DEFINED__
#undef __DiagramNodeChildren_FWD_DEFINED__
#undef __DiagramNodes_FWD_DEFINED__
#undef __Diagram_FWD_DEFINED__
#undef __CustomProperty_FWD_DEFINED__
#undef __CustomProperties_FWD_DEFINED__
#undef __SmartTag_FWD_DEFINED__
#undef __SmartTags_FWD_DEFINED__
#undef __StyleSheet_FWD_DEFINED__
#undef __StyleSheets_FWD_DEFINED__
#undef __MappedDataField_FWD_DEFINED__
#undef __MappedDataFields_FWD_DEFINED__
#undef __CanvasShapes_FWD_DEFINED__
#undef __TableStyle_FWD_DEFINED__
#undef __ConditionalStyle_FWD_DEFINED__
#undef __FootnoteOptions_FWD_DEFINED__
#undef __EndnoteOptions_FWD_DEFINED__
#undef __Reviewers_FWD_DEFINED__
#undef __Reviewer_FWD_DEFINED__
#undef __TaskPane_FWD_DEFINED__
#undef __TaskPanes_FWD_DEFINED__
#undef __IApplicationEvents3_FWD_DEFINED__
#undef __ApplicationEvents3_FWD_DEFINED__
#undef __SmartTagAction_FWD_DEFINED__
#undef __SmartTagActions_FWD_DEFINED__
#undef __SmartTagRecognizer_FWD_DEFINED__
#undef __SmartTagRecognizers_FWD_DEFINED__
#undef __SmartTagType_FWD_DEFINED__
#undef __SmartTagTypes_FWD_DEFINED__
#undef __Line_FWD_DEFINED__
#undef __Lines_FWD_DEFINED__
#undef __Rectangle_FWD_DEFINED__
#undef __Rectangles_FWD_DEFINED__
#undef __Break_FWD_DEFINED__
#undef __Breaks_FWD_DEFINED__
#undef __Page_FWD_DEFINED__
#undef __Pages_FWD_DEFINED__
#undef __XMLNode_FWD_DEFINED__
#undef __XMLNodes_FWD_DEFINED__
#undef __XMLSchemaReference_FWD_DEFINED__
#undef __XMLSchemaReferences_FWD_DEFINED__
#undef __XMLChildNodeSuggestion_FWD_DEFINED__
#undef __XMLChildNodeSuggestions_FWD_DEFINED__
#undef __XMLNamespace_FWD_DEFINED__
#undef __XMLNamespaces_FWD_DEFINED__
#undef __XSLTransform_FWD_DEFINED__
#undef __XSLTransforms_FWD_DEFINED__
#undef __Editors_FWD_DEFINED__
#undef __Editor_FWD_DEFINED__
#undef __IApplicationEvents4_FWD_DEFINED__
#undef __ApplicationEvents4_FWD_DEFINED__
#undef __DocumentEvents2_FWD_DEFINED__
#undef __Source_FWD_DEFINED__
#undef __Sources_FWD_DEFINED__
#undef __Bibliography_FWD_DEFINED__
#undef __OMaths_FWD_DEFINED__
#undef __OMath_FWD_DEFINED__
#undef __OMathFunctions_FWD_DEFINED__
#undef __OMathArgs_FWD_DEFINED__
#undef __OMathFunction_FWD_DEFINED__
#undef __OMathAcc_FWD_DEFINED__
#undef __OMathBar_FWD_DEFINED__
#undef __OMathBox_FWD_DEFINED__
#undef __OMathBorderBox_FWD_DEFINED__
#undef __OMathDelim_FWD_DEFINED__
#undef __OMathEqArray_FWD_DEFINED__
#undef __OMathFrac_FWD_DEFINED__
#undef __OMathFunc_FWD_DEFINED__
#undef __OMathGroupChar_FWD_DEFINED__
#undef __OMathLimLow_FWD_DEFINED__
#undef __OMathLimUpp_FWD_DEFINED__
#undef __OMathMat_FWD_DEFINED__
#undef __OMathMatRows_FWD_DEFINED__
#undef __OMathMatCols_FWD_DEFINED__
#undef __OMathMatRow_FWD_DEFINED__
#undef __OMathMatCol_FWD_DEFINED__
#undef __OMathNary_FWD_DEFINED__
#undef __OMathPhantom_FWD_DEFINED__
#undef __OMathScrPre_FWD_DEFINED__
#undef __OMathRad_FWD_DEFINED__
#undef __OMathScrSub_FWD_DEFINED__
#undef __OMathScrSubSup_FWD_DEFINED__
#undef __OMathScrSup_FWD_DEFINED__
#undef __OMathAutoCorrect_FWD_DEFINED__
#undef __OMathAutoCorrectEntries_FWD_DEFINED__
#undef __OMathAutoCorrectEntry_FWD_DEFINED__
#undef __OMathRecognizedFunctions_FWD_DEFINED__
#undef __OMathRecognizedFunction_FWD_DEFINED__
#undef __ContentControls_FWD_DEFINED__
#undef __ContentControl_FWD_DEFINED__
#undef __XMLMapping_FWD_DEFINED__
#undef __ContentControlListEntries_FWD_DEFINED__
#undef __ContentControlListEntry_FWD_DEFINED__
#undef __BuildingBlockTypes_FWD_DEFINED__
#undef __BuildingBlockType_FWD_DEFINED__
#undef __Categories_FWD_DEFINED__
#undef __Category_FWD_DEFINED__
#undef __BuildingBlocks_FWD_DEFINED__
#undef __BuildingBlock_FWD_DEFINED__
#undef __BuildingBlockEntries_FWD_DEFINED__
#undef __OMathBreaks_FWD_DEFINED__
#undef __OMathBreak_FWD_DEFINED__
#undef __Research_FWD_DEFINED__
#undef __SoftEdgeFormat_FWD_DEFINED__
#undef __GlowFormat_FWD_DEFINED__
#undef __ReflectionFormat_FWD_DEFINED__
#undef __ChartData_FWD_DEFINED__
#undef __Chart_FWD_DEFINED__
#undef __Corners_FWD_DEFINED__
#undef __Legend_FWD_DEFINED__
#undef __ChartBorder_FWD_DEFINED__
#undef __Walls_FWD_DEFINED__
#undef __Floor_FWD_DEFINED__
#undef __PlotArea_FWD_DEFINED__
#undef __ChartArea_FWD_DEFINED__
#undef __SeriesLines_FWD_DEFINED__
#undef __LeaderLines_FWD_DEFINED__
#undef __Gridlines_FWD_DEFINED__
#undef __UpBars_FWD_DEFINED__
#undef __DownBars_FWD_DEFINED__
#undef __Interior_FWD_DEFINED__
#undef __ChartFillFormat_FWD_DEFINED__
#undef __LegendEntries_FWD_DEFINED__
#undef __ChartFont_FWD_DEFINED__
#undef __ChartColorFormat_FWD_DEFINED__
#undef __LegendEntry_FWD_DEFINED__
#undef __LegendKey_FWD_DEFINED__
#undef __SeriesCollection_FWD_DEFINED__
#undef __Series_FWD_DEFINED__
#undef __ErrorBars_FWD_DEFINED__
#undef __Trendline_FWD_DEFINED__
#undef __Trendlines_FWD_DEFINED__
#undef __DataLabels_FWD_DEFINED__
#undef __DataLabel_FWD_DEFINED__
#undef __Points_FWD_DEFINED__
#undef __Point_FWD_DEFINED__
#undef __Axes_FWD_DEFINED__
#undef __Axis_FWD_DEFINED__
#undef __DataTable_FWD_DEFINED__
#undef __ChartTitle_FWD_DEFINED__
#undef __AxisTitle_FWD_DEFINED__
#undef __DisplayUnitLabel_FWD_DEFINED__
#undef __TickLabels_FWD_DEFINED__
#undef __DropLines_FWD_DEFINED__
#undef __HiLoLines_FWD_DEFINED__
#undef __ChartGroup_FWD_DEFINED__
#undef __ChartGroups_FWD_DEFINED__
#undef __ChartCharacters_FWD_DEFINED__
#undef __ChartFormat_FWD_DEFINED__
#undef __UndoRecord_FWD_DEFINED__
#undef __CoAuthLock_FWD_DEFINED__
#undef __CoAuthLocks_FWD_DEFINED__
#undef __CoAuthUpdate_FWD_DEFINED__
#undef __CoAuthUpdates_FWD_DEFINED__
#undef __CoAuthor_FWD_DEFINED__
#undef __CoAuthors_FWD_DEFINED__
#undef __CoAuthoring_FWD_DEFINED__
#undef __Conflicts_FWD_DEFINED__
#undef __Conflict_FWD_DEFINED__
#undef __ProtectedViewWindows_FWD_DEFINED__
#undef __ProtectedViewWindow_FWD_DEFINED__
#undef __RepeatingSectionItemColl_FWD_DEFINED__
#undef __RepeatingSectionItem_FWD_DEFINED__
#undef __FullSeriesCollection_FWD_DEFINED__
#undef __ChartCategory_FWD_DEFINED__
#undef __CategoryCollection_FWD_DEFINED__
#undef __Broadcast_FWD_DEFINED__
#undef __RevisionsFilter_FWD_DEFINED__
#undef __DocumentField_FWD_DEFINED__
#undef __DocumentFields_FWD_DEFINED__
#undef __Global_FWD_DEFINED__
#undef __Application_FWD_DEFINED__
#undef __Document_FWD_DEFINED__
#undef __OLEControl_FWD_DEFINED__
#undef __Font_FWD_DEFINED__
#undef __ParagraphFormat_FWD_DEFINED__
#undef __LetterContent_FWD_DEFINED__
#undef __IWORDCtrlExtender_FWD_DEFINED__
#undef __KWORDCtrlExtender_FWD_DEFINED__
#undef __Word_LIBRARY_DEFINED__
#undef ___Application_INTERFACE_DEFINED__
#undef ___Global_INTERFACE_DEFINED__
#undef __FontNames_INTERFACE_DEFINED__
#undef __Languages_INTERFACE_DEFINED__
#undef __Language_INTERFACE_DEFINED__
#undef __Documents_INTERFACE_DEFINED__
#undef ___Document_INTERFACE_DEFINED__
#undef __Template_INTERFACE_DEFINED__
#undef __Templates_INTERFACE_DEFINED__
#undef __RoutingSlip_INTERFACE_DEFINED__
#undef __Bookmark_INTERFACE_DEFINED__
#undef __Bookmarks_INTERFACE_DEFINED__
#undef __Variable_INTERFACE_DEFINED__
#undef __Variables_INTERFACE_DEFINED__
#undef __RecentFile_INTERFACE_DEFINED__
#undef __RecentFiles_INTERFACE_DEFINED__
#undef __Window_INTERFACE_DEFINED__
#undef __Windows_INTERFACE_DEFINED__
#undef __Pane_INTERFACE_DEFINED__
#undef __Panes_INTERFACE_DEFINED__
#undef __Range_INTERFACE_DEFINED__
#undef __ListFormat_INTERFACE_DEFINED__
#undef __Find_INTERFACE_DEFINED__
#undef __Replacement_INTERFACE_DEFINED__
#undef __Characters_INTERFACE_DEFINED__
#undef __Words_INTERFACE_DEFINED__
#undef __Sentences_INTERFACE_DEFINED__
#undef __Sections_INTERFACE_DEFINED__
#undef __Section_INTERFACE_DEFINED__
#undef __Paragraphs_INTERFACE_DEFINED__
#undef __Paragraph_INTERFACE_DEFINED__
#undef __DropCap_INTERFACE_DEFINED__
#undef __TabStops_INTERFACE_DEFINED__
#undef __TabStop_INTERFACE_DEFINED__
#undef ___ParagraphFormat_INTERFACE_DEFINED__
#undef ___Font_INTERFACE_DEFINED__
#undef __Table_INTERFACE_DEFINED__
#undef __Row_INTERFACE_DEFINED__
#undef __Column_INTERFACE_DEFINED__
#undef __Cell_INTERFACE_DEFINED__
#undef __Tables_INTERFACE_DEFINED__
#undef __Rows_INTERFACE_DEFINED__
#undef __Columns_INTERFACE_DEFINED__
#undef __Cells_INTERFACE_DEFINED__
#undef __AutoCorrect_INTERFACE_DEFINED__
#undef __AutoCorrectEntries_INTERFACE_DEFINED__
#undef __AutoCorrectEntry_INTERFACE_DEFINED__
#undef __FirstLetterExceptions_INTERFACE_DEFINED__
#undef __FirstLetterException_INTERFACE_DEFINED__
#undef __TwoInitialCapsExceptions_INTERFACE_DEFINED__
#undef __TwoInitialCapsException_INTERFACE_DEFINED__
#undef __Footnotes_INTERFACE_DEFINED__
#undef __Endnotes_INTERFACE_DEFINED__
#undef __Comments_INTERFACE_DEFINED__
#undef __Footnote_INTERFACE_DEFINED__
#undef __Endnote_INTERFACE_DEFINED__
#undef __Comment_INTERFACE_DEFINED__
#undef __Borders_INTERFACE_DEFINED__
#undef __Border_INTERFACE_DEFINED__
#undef __Shading_INTERFACE_DEFINED__
#undef __TextRetrievalMode_INTERFACE_DEFINED__
#undef __AutoTextEntries_INTERFACE_DEFINED__
#undef __AutoTextEntry_INTERFACE_DEFINED__
#undef __System_INTERFACE_DEFINED__
#undef __OLEFormat_INTERFACE_DEFINED__
#undef __LinkFormat_INTERFACE_DEFINED__
#undef ___OLEControl_INTERFACE_DEFINED__
#undef __Fields_INTERFACE_DEFINED__
#undef __Field_INTERFACE_DEFINED__
#undef __Browser_INTERFACE_DEFINED__
#undef __Styles_INTERFACE_DEFINED__
#undef __Style_INTERFACE_DEFINED__
#undef __Frames_INTERFACE_DEFINED__
#undef __Frame_INTERFACE_DEFINED__
#undef __FormFields_INTERFACE_DEFINED__
#undef __FormField_INTERFACE_DEFINED__
#undef __TextInput_INTERFACE_DEFINED__
#undef __CheckBox_INTERFACE_DEFINED__
#undef __DropDown_INTERFACE_DEFINED__
#undef __ListEntries_INTERFACE_DEFINED__
#undef __ListEntry_INTERFACE_DEFINED__
#undef __TablesOfFigures_INTERFACE_DEFINED__
#undef __TableOfFigures_INTERFACE_DEFINED__
#undef __MailMerge_INTERFACE_DEFINED__
#undef __MailMergeFields_INTERFACE_DEFINED__
#undef __MailMergeField_INTERFACE_DEFINED__
#undef __MailMergeDataSource_INTERFACE_DEFINED__
#undef __MailMergeFieldNames_INTERFACE_DEFINED__
#undef __MailMergeFieldName_INTERFACE_DEFINED__
#undef __MailMergeDataFields_INTERFACE_DEFINED__
#undef __MailMergeDataField_INTERFACE_DEFINED__
#undef __Envelope_INTERFACE_DEFINED__
#undef __MailingLabel_INTERFACE_DEFINED__
#undef __CustomLabels_INTERFACE_DEFINED__
#undef __CustomLabel_INTERFACE_DEFINED__
#undef __TablesOfContents_INTERFACE_DEFINED__
#undef __TableOfContents_INTERFACE_DEFINED__
#undef __TablesOfAuthorities_INTERFACE_DEFINED__
#undef __TableOfAuthorities_INTERFACE_DEFINED__
#undef __Dialogs_INTERFACE_DEFINED__
#undef __Dialog_INTERFACE_DEFINED__
#undef __PageSetup_INTERFACE_DEFINED__
#undef __LineNumbering_INTERFACE_DEFINED__
#undef __TextColumns_INTERFACE_DEFINED__
#undef __TextColumn_INTERFACE_DEFINED__
#undef __Selection_INTERFACE_DEFINED__
#undef __TablesOfAuthoritiesCategories_INTERFACE_DEFINED__
#undef __TableOfAuthoritiesCategory_INTERFACE_DEFINED__
#undef __CaptionLabels_INTERFACE_DEFINED__
#undef __CaptionLabel_INTERFACE_DEFINED__
#undef __AutoCaptions_INTERFACE_DEFINED__
#undef __AutoCaption_INTERFACE_DEFINED__
#undef __Indexes_INTERFACE_DEFINED__
#undef __Index_INTERFACE_DEFINED__
#undef __AddIn_INTERFACE_DEFINED__
#undef __AddIns_INTERFACE_DEFINED__
#undef __Revisions_INTERFACE_DEFINED__
#undef __Revision_INTERFACE_DEFINED__
#undef __Task_INTERFACE_DEFINED__
#undef __Tasks_INTERFACE_DEFINED__
#undef __HeadersFooters_INTERFACE_DEFINED__
#undef __HeaderFooter_INTERFACE_DEFINED__
#undef __PageNumbers_INTERFACE_DEFINED__
#undef __PageNumber_INTERFACE_DEFINED__
#undef __Subdocuments_INTERFACE_DEFINED__
#undef __Subdocument_INTERFACE_DEFINED__
#undef __HeadingStyles_INTERFACE_DEFINED__
#undef __HeadingStyle_INTERFACE_DEFINED__
#undef __StoryRanges_INTERFACE_DEFINED__
#undef __ListLevel_INTERFACE_DEFINED__
#undef __ListLevels_INTERFACE_DEFINED__
#undef __ListTemplate_INTERFACE_DEFINED__
#undef __ListTemplates_INTERFACE_DEFINED__
#undef __ListParagraphs_INTERFACE_DEFINED__
#undef __List_INTERFACE_DEFINED__
#undef __Lists_INTERFACE_DEFINED__
#undef __ListGallery_INTERFACE_DEFINED__
#undef __ListGalleries_INTERFACE_DEFINED__
#undef __KeyBindings_INTERFACE_DEFINED__
#undef __KeysBoundTo_INTERFACE_DEFINED__
#undef __KeyBinding_INTERFACE_DEFINED__
#undef __FileConverter_INTERFACE_DEFINED__
#undef __FileConverters_INTERFACE_DEFINED__
#undef __SynonymInfo_INTERFACE_DEFINED__
#undef __Hyperlinks_INTERFACE_DEFINED__
#undef __Hyperlink_INTERFACE_DEFINED__
#undef __Shapes_INTERFACE_DEFINED__
#undef __ShapeRange_INTERFACE_DEFINED__
#undef __GroupShapes_INTERFACE_DEFINED__
#undef __Shape_INTERFACE_DEFINED__
#undef __TextFrame_INTERFACE_DEFINED__
#undef ___LetterContent_INTERFACE_DEFINED__
#undef __View_INTERFACE_DEFINED__
#undef __Zoom_INTERFACE_DEFINED__
#undef __Zooms_INTERFACE_DEFINED__
#undef __InlineShape_INTERFACE_DEFINED__
#undef __InlineShapes_INTERFACE_DEFINED__
#undef __SpellingSuggestions_INTERFACE_DEFINED__
#undef __SpellingSuggestion_INTERFACE_DEFINED__
#undef __Dictionaries_INTERFACE_DEFINED__
#undef __HangulHanjaConversionDictionaries_INTERFACE_DEFINED__
#undef __Dictionary_INTERFACE_DEFINED__
#undef __ReadabilityStatistics_INTERFACE_DEFINED__
#undef __ReadabilityStatistic_INTERFACE_DEFINED__
#undef __Versions_INTERFACE_DEFINED__
#undef __Version_INTERFACE_DEFINED__
#undef __Options_INTERFACE_DEFINED__
#undef __MailMessage_INTERFACE_DEFINED__
#undef __ProofreadingErrors_INTERFACE_DEFINED__
#undef __Mailer_INTERFACE_DEFINED__
#undef __WrapFormat_INTERFACE_DEFINED__
#undef __HangulAndAlphabetExceptions_INTERFACE_DEFINED__
#undef __HangulAndAlphabetException_INTERFACE_DEFINED__
#undef __Adjustments_INTERFACE_DEFINED__
#undef __CalloutFormat_INTERFACE_DEFINED__
#undef __ColorFormat_INTERFACE_DEFINED__
#undef __ConnectorFormat_INTERFACE_DEFINED__
#undef __FillFormat_INTERFACE_DEFINED__
#undef __FreeformBuilder_INTERFACE_DEFINED__
#undef __LineFormat_INTERFACE_DEFINED__
#undef __PictureFormat_INTERFACE_DEFINED__
#undef __ShadowFormat_INTERFACE_DEFINED__
#undef __ShapeNode_INTERFACE_DEFINED__
#undef __ShapeNodes_INTERFACE_DEFINED__
#undef __TextEffectFormat_INTERFACE_DEFINED__
#undef __ThreeDFormat_INTERFACE_DEFINED__
#undef __ApplicationEvents_DISPINTERFACE_DEFINED__
#undef __DocumentEvents_DISPINTERFACE_DEFINED__
#undef __OCXEvents_DISPINTERFACE_DEFINED__
#undef __IApplicationEvents_INTERFACE_DEFINED__
#undef __IApplicationEvents2_INTERFACE_DEFINED__
#undef __ApplicationEvents2_DISPINTERFACE_DEFINED__
#undef __EmailAuthor_INTERFACE_DEFINED__
#undef __EmailOptions_INTERFACE_DEFINED__
#undef __EmailSignature_INTERFACE_DEFINED__
#undef __Email_INTERFACE_DEFINED__
#undef __HorizontalLineFormat_INTERFACE_DEFINED__
#undef __Frameset_INTERFACE_DEFINED__
#undef __DefaultWebOptions_INTERFACE_DEFINED__
#undef __WebOptions_INTERFACE_DEFINED__
#undef __OtherCorrectionsExceptions_INTERFACE_DEFINED__
#undef __OtherCorrectionsException_INTERFACE_DEFINED__
#undef __EmailSignatureEntries_INTERFACE_DEFINED__
#undef __EmailSignatureEntry_INTERFACE_DEFINED__
#undef __HTMLDivision_INTERFACE_DEFINED__
#undef __HTMLDivisions_INTERFACE_DEFINED__
#undef __DiagramNode_INTERFACE_DEFINED__
#undef __DiagramNodeChildren_INTERFACE_DEFINED__
#undef __DiagramNodes_INTERFACE_DEFINED__
#undef __Diagram_INTERFACE_DEFINED__
#undef __CustomProperty_INTERFACE_DEFINED__
#undef __CustomProperties_INTERFACE_DEFINED__
#undef __SmartTag_INTERFACE_DEFINED__
#undef __SmartTags_INTERFACE_DEFINED__
#undef __StyleSheet_INTERFACE_DEFINED__
#undef __StyleSheets_INTERFACE_DEFINED__
#undef __MappedDataField_INTERFACE_DEFINED__
#undef __MappedDataFields_INTERFACE_DEFINED__
#undef __CanvasShapes_INTERFACE_DEFINED__
#undef __TableStyle_INTERFACE_DEFINED__
#undef __ConditionalStyle_INTERFACE_DEFINED__
#undef __FootnoteOptions_INTERFACE_DEFINED__
#undef __EndnoteOptions_INTERFACE_DEFINED__
#undef __Reviewers_INTERFACE_DEFINED__
#undef __Reviewer_INTERFACE_DEFINED__
#undef __TaskPane_INTERFACE_DEFINED__
#undef __TaskPanes_INTERFACE_DEFINED__
#undef __IApplicationEvents3_INTERFACE_DEFINED__
#undef __ApplicationEvents3_DISPINTERFACE_DEFINED__
#undef __SmartTagAction_INTERFACE_DEFINED__
#undef __SmartTagActions_INTERFACE_DEFINED__
#undef __SmartTagRecognizer_INTERFACE_DEFINED__
#undef __SmartTagRecognizers_INTERFACE_DEFINED__
#undef __SmartTagType_INTERFACE_DEFINED__
#undef __SmartTagTypes_INTERFACE_DEFINED__
#undef __Line_INTERFACE_DEFINED__
#undef __Lines_INTERFACE_DEFINED__
#undef __Rectangle_INTERFACE_DEFINED__
#undef __Rectangles_INTERFACE_DEFINED__
#undef __Break_INTERFACE_DEFINED__
#undef __Breaks_INTERFACE_DEFINED__
#undef __Page_INTERFACE_DEFINED__
#undef __Pages_INTERFACE_DEFINED__
#undef __XMLNode_INTERFACE_DEFINED__
#undef __XMLNodes_INTERFACE_DEFINED__
#undef __XMLSchemaReference_INTERFACE_DEFINED__
#undef __XMLSchemaReferences_INTERFACE_DEFINED__
#undef __XMLChildNodeSuggestion_INTERFACE_DEFINED__
#undef __XMLChildNodeSuggestions_INTERFACE_DEFINED__
#undef __XMLNamespace_INTERFACE_DEFINED__
#undef __XMLNamespaces_INTERFACE_DEFINED__
#undef __XSLTransform_INTERFACE_DEFINED__
#undef __XSLTransforms_INTERFACE_DEFINED__
#undef __Editors_INTERFACE_DEFINED__
#undef __Editor_INTERFACE_DEFINED__
#undef __IApplicationEvents4_INTERFACE_DEFINED__
#undef __ApplicationEvents4_DISPINTERFACE_DEFINED__
#undef __DocumentEvents2_DISPINTERFACE_DEFINED__
#undef __Source_INTERFACE_DEFINED__
#undef __Sources_INTERFACE_DEFINED__
#undef __Bibliography_INTERFACE_DEFINED__
#undef __OMaths_INTERFACE_DEFINED__
#undef __OMath_INTERFACE_DEFINED__
#undef __OMathFunctions_INTERFACE_DEFINED__
#undef __OMathArgs_INTERFACE_DEFINED__
#undef __OMathFunction_INTERFACE_DEFINED__
#undef __OMathAcc_INTERFACE_DEFINED__
#undef __OMathBar_INTERFACE_DEFINED__
#undef __OMathBox_INTERFACE_DEFINED__
#undef __OMathBorderBox_INTERFACE_DEFINED__
#undef __OMathDelim_INTERFACE_DEFINED__
#undef __OMathEqArray_INTERFACE_DEFINED__
#undef __OMathFrac_INTERFACE_DEFINED__
#undef __OMathFunc_INTERFACE_DEFINED__
#undef __OMathGroupChar_INTERFACE_DEFINED__
#undef __OMathLimLow_INTERFACE_DEFINED__
#undef __OMathLimUpp_INTERFACE_DEFINED__
#undef __OMathMat_INTERFACE_DEFINED__
#undef __OMathMatRows_INTERFACE_DEFINED__
#undef __OMathMatCols_INTERFACE_DEFINED__
#undef __OMathMatRow_INTERFACE_DEFINED__
#undef __OMathMatCol_INTERFACE_DEFINED__
#undef __OMathNary_INTERFACE_DEFINED__
#undef __OMathPhantom_INTERFACE_DEFINED__
#undef __OMathScrPre_INTERFACE_DEFINED__
#undef __OMathRad_INTERFACE_DEFINED__
#undef __OMathScrSub_INTERFACE_DEFINED__
#undef __OMathScrSubSup_INTERFACE_DEFINED__
#undef __OMathScrSup_INTERFACE_DEFINED__
#undef __OMathAutoCorrect_INTERFACE_DEFINED__
#undef __OMathAutoCorrectEntries_INTERFACE_DEFINED__
#undef __OMathAutoCorrectEntry_INTERFACE_DEFINED__
#undef __OMathRecognizedFunctions_INTERFACE_DEFINED__
#undef __OMathRecognizedFunction_INTERFACE_DEFINED__
#undef __ContentControls_INTERFACE_DEFINED__
#undef __ContentControl_INTERFACE_DEFINED__
#undef __XMLMapping_INTERFACE_DEFINED__
#undef __ContentControlListEntries_INTERFACE_DEFINED__
#undef __ContentControlListEntry_INTERFACE_DEFINED__
#undef __BuildingBlockTypes_INTERFACE_DEFINED__
#undef __BuildingBlockType_INTERFACE_DEFINED__
#undef __Categories_INTERFACE_DEFINED__
#undef __Category_INTERFACE_DEFINED__
#undef __BuildingBlocks_INTERFACE_DEFINED__
#undef __BuildingBlock_INTERFACE_DEFINED__
#undef __BuildingBlockEntries_INTERFACE_DEFINED__
#undef __OMathBreaks_INTERFACE_DEFINED__
#undef __OMathBreak_INTERFACE_DEFINED__
#undef __Research_INTERFACE_DEFINED__
#undef __SoftEdgeFormat_INTERFACE_DEFINED__
#undef __GlowFormat_INTERFACE_DEFINED__
#undef __ReflectionFormat_INTERFACE_DEFINED__
#undef __ChartData_INTERFACE_DEFINED__
#undef __Chart_INTERFACE_DEFINED__
#undef __Corners_INTERFACE_DEFINED__
#undef __Legend_INTERFACE_DEFINED__
#undef __ChartBorder_INTERFACE_DEFINED__
#undef __Walls_INTERFACE_DEFINED__
#undef __Floor_INTERFACE_DEFINED__
#undef __PlotArea_INTERFACE_DEFINED__
#undef __ChartArea_INTERFACE_DEFINED__
#undef __SeriesLines_INTERFACE_DEFINED__
#undef __LeaderLines_INTERFACE_DEFINED__
#undef __Gridlines_INTERFACE_DEFINED__
#undef __UpBars_INTERFACE_DEFINED__
#undef __DownBars_INTERFACE_DEFINED__
#undef __Interior_INTERFACE_DEFINED__
#undef __ChartFillFormat_INTERFACE_DEFINED__
#undef __LegendEntries_INTERFACE_DEFINED__
#undef __ChartFont_INTERFACE_DEFINED__
#undef __ChartColorFormat_INTERFACE_DEFINED__
#undef __LegendEntry_INTERFACE_DEFINED__
#undef __LegendKey_INTERFACE_DEFINED__
#undef __SeriesCollection_INTERFACE_DEFINED__
#undef __Series_INTERFACE_DEFINED__
#undef __ErrorBars_INTERFACE_DEFINED__
#undef __Trendline_INTERFACE_DEFINED__
#undef __Trendlines_INTERFACE_DEFINED__
#undef __DataLabels_INTERFACE_DEFINED__
#undef __DataLabel_INTERFACE_DEFINED__
#undef __Points_INTERFACE_DEFINED__
#undef __Point_INTERFACE_DEFINED__
#undef __Axes_INTERFACE_DEFINED__
#undef __Axis_INTERFACE_DEFINED__
#undef __DataTable_INTERFACE_DEFINED__
#undef __ChartTitle_INTERFACE_DEFINED__
#undef __AxisTitle_INTERFACE_DEFINED__
#undef __DisplayUnitLabel_INTERFACE_DEFINED__
#undef __TickLabels_INTERFACE_DEFINED__
#undef __DropLines_INTERFACE_DEFINED__
#undef __HiLoLines_INTERFACE_DEFINED__
#undef __ChartGroup_INTERFACE_DEFINED__
#undef __ChartGroups_INTERFACE_DEFINED__
#undef __ChartCharacters_INTERFACE_DEFINED__
#undef __ChartFormat_INTERFACE_DEFINED__
#undef __UndoRecord_INTERFACE_DEFINED__
#undef __CoAuthLock_INTERFACE_DEFINED__
#undef __CoAuthLocks_INTERFACE_DEFINED__
#undef __CoAuthUpdate_INTERFACE_DEFINED__
#undef __CoAuthUpdates_INTERFACE_DEFINED__
#undef __CoAuthor_INTERFACE_DEFINED__
#undef __CoAuthors_INTERFACE_DEFINED__
#undef __CoAuthoring_INTERFACE_DEFINED__
#undef __Conflicts_INTERFACE_DEFINED__
#undef __Conflict_INTERFACE_DEFINED__
#undef __ProtectedViewWindows_INTERFACE_DEFINED__
#undef __ProtectedViewWindow_INTERFACE_DEFINED__
#undef __RepeatingSectionItemColl_INTERFACE_DEFINED__
#undef __RepeatingSectionItem_INTERFACE_DEFINED__
#undef __FullSeriesCollection_INTERFACE_DEFINED__
#undef __ChartCategory_INTERFACE_DEFINED__
#undef __CategoryCollection_INTERFACE_DEFINED__
#undef __Broadcast_INTERFACE_DEFINED__
#undef __RevisionsFilter_INTERFACE_DEFINED__
#undef __DocumentField_INTERFACE_DEFINED__
#undef __DocumentFields_INTERFACE_DEFINED__
#undef __IWORDCtrlExtender_INTERFACE_DEFINED__

// INCLUDE MSO ENUM
#undef __INCLUDE_MSO_ENUM_H__
#undef __INCLUDE_MSO_ENUM_CHART_H__
