<!--- 
License:
pptutils:
Copyright 2007 Todd Sharp
  
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
 --->

<cfcomponent displayname="ppt utils" output="false">
	
	<cfset instance.paths = arrayNew(1) />
	
	<cfset instance.qLibDir = getLib() />
	
	<cfloop query="instance.qLibDir">
	   <cfset arrayAppend(instance.paths, "#instance.qLibDir.directory#\#instance.qLibDir.name#")>
	</cfloop>

	<cfset instance.loader = createObject("component", "javaloader.JavaLoader").init(instance.paths) />

	<cffunction name="init" access="public" returntype="pptutils" output="false">
		<cfreturn this />
	</cffunction>
	
	<cffunction name="getLib" access="private" output="false" returntype="query" hint="i read the lib dir and return the files">
		<cfset var libDir = getDirectoryFromPath(getCurrentTemplatePath()) & "lib" />
		<cfset var qLibDir = queryNew("") />
		<cfdirectory action="list" directory="#libDir#" name="qLibDir" />
		<cfreturn qLibDir />
	</cffunction>
	
	<cffunction name="extractText" access="public" returntype="array" output="false" hint="i extract text from a PPT by means of an array of structs containing an array element for each slide in the PowerPoint">
		<cfargument name="pathToPPT" required="true" hint="the full path to the powerpoint to convert" />
		<cfset var hslf = instance.loader.create("org.apache.poi.hslf.HSLFSlideShow").init(arguments.pathToPPT) />
		<cfset var slideshow = instance.loader.create("org.apache.poi.hslf.usermodel.SlideShow").init(hslf) />
		<cfset var slides = slideshow.getSlides() />
		<cfset var retArr = arrayNew(1) />
		<cfset var slide = structNew() />
		<cfset var i = "" />
		<cfset var j = "" />
		<cfset var k = "" />
		<cfset var thisSlide = "" />
		<cfset var thisSlideText = "" />
		<cfset var thisSlideRichText = "" />
		<cfset var rawText = "" />
		<cfset var slideText = "" />

		<cfloop from="1" to="#arrayLen(slides)#" index="i">
			<cfset slide.slideText = structNew() />
			<cfset thisSlide = slides[i] />
			<cfset slide.slideTitle = thisSlide.getTitle() />	
			<cfset thisSlideText = thisSlide.getTextRuns() />
			<cfset slideText = "" />
			
			<cfloop from="1" to="#arrayLen(thisSlideText)#" index="j">
				<cfset thisSlideRichText = thisSlideText[j].getRichTextRuns() />
				<cfloop from="1" to="#arrayLen(thisSlideRichText)#" index="k">
					<cfset rawText = thisSlideRichText[k].getText() />		
					<cfset slideText = slideText & rawText />	
				</cfloop>
			</cfloop>
			
			<cfset slide.slideText = duplicate(slideText) />
			<cfset arrayAppend(retArr, duplicate(slide)) />
			
		</cfloop>
		
		<cfreturn retArr />
	</cffunction>
	
	<!--- convertPowerPoint is depracated - this method will call readPowerPoint() --->
	<cffunction name="convertPowerPoint" access="public" output="false" returntype="array" hint="convertPowerPoint method is depracated.  use readPowerPoint() instead.">
		<cfargument name="pathToPPT" required="true" hint="the full path to the powerpoint to convert" />
		<cfreturn readPowerPoint(arguments.pathToPPT) />
	</cffunction>
	
	<cffunction name="readPowerPoint" access="public" output="true" returntype="array" hint="i read a powerpoint and return an array.  each array element represents a slide in the slideshow.">
		<cfargument name="pathToPPT" required="true" hint="the full path to the powerpoint to convert" />
		<cfset var hslf = instance.loader.create("org.apache.poi.hslf.HSLFSlideShow").init(javacast("string", arguments.pathToPPT)) />
		<cfset var slideshow = instance.loader.create("org.apache.poi.hslf.usermodel.SlideShow").init(hslf) />
		<cfset var slides = slideshow.getSlides() />
		<cfset var slideBgColor = "" />
		<cfset var thisSlide = "" />
		<cfset var thisSlideShape = "" />
		<cfset var thisSlideShapes = "" />
		<cfset var thisSlideDimensions = "" />
		<cfset var slideStruct = structNew() />
		<cfset var i = "" />
		<cfset var retArr = arrayNew(1) />
		<cfset var shapeClass = "" />
		<cfset var anchor = "" />
		<cfset var textBox = structNew() />
		<cfset var line = structNew() />
		<cfset var shape = structNew() />
		<cfset var img = structNew() />
		<cfset var slideTextSpans = "" />
		<cfset var thisTextBoxTextSpans = arrayNew(1) />
		<cfset var textSpan = structNew() />
		<cfset var j = "" />
		<cfset var thisSlideNotes = "" />
		<cfset var k = "" />
		<cfset var slideNotes = structNew() />
		<cfset var slideBgImage = "" />
		
		<cfloop from="1" to="#arrayLen(slides)#" index="i">
			<cfset slideStruct = structNew() />
			<cfset thisSlide = slides[i] />
			<cfset thisSlideShapes = thisSlide.getShapes() />
			<cfset thisSlideDimensions = slideshow.getPageSize() />
			<cfset slideStruct.slideTitle = thisSlide.getTitle() />
			<cfset slideStruct.slideWidth = thisSlideDimensions.width />
			<cfset slideStruct.slideHeight = thisSlideDimensions.height />
			<cfset slideStruct.slideBackgroundImage = structNew() /> 
			<cftry>
				<cfset slideStruct.slideBackgroundImage.imgData = fixNull(thisSlide.getBackground().getFill().getPictureData().getData()) />
				<cfset slideStruct.slideBackgroundImage.imgType = fixNull(getImageType(thisSlide.getBackground().getFill().getPictureData().getType())) />
				<cfcatch>
					<cfset slideStruct.slideBackgroundImage.imgData = "" />
					<cfset slideStruct.slideBackgroundImage.imgType = "" />				
				</cfcatch>
			</cftry>
			<!--- if no background the var is undefined - set it to an empty string --->
			
			<cftry>
				<cfset slideBgColor = thisSlide.getBackground().getFill().getBackgroundColor() />
				<cfset slideStruct.slideBackgroundColor = slideBgColor.getRed() & "," & slideBgColor.getGreen() & "," & slideBgColor.getBlue() />
				<cfcatch>
					<cfset slideStruct.slideBackgroundColor = "" />
				</cfcatch>
			</cftry>

			<cfset slideStruct.notes = arrayNew(1) />
			<cftry>
				<cfset thisSlideNotes = fixNull(thisSlide.getNotesSheet().getTextRuns()) />
	
				<cfloop from="1" to="#arrayLen(thisSlideNotes)#" index="k">
					<cfset slideNotes = structNew() />
					<cfset slideNotes.noteText = thisSlideNotes[k].getText() />
					<cfset arrayAppend(slideStruct.notes, slideNotes) />
				</cfloop>
				<cfcatch type="any">
					<cfset slideNotes = structNew() />
					<cfset slideNotes.noteText = "" />
					<cfset arrayAppend(slideStruct.notes, slideNotes) />
				</cfcatch>
			</cftry>
			<cfset slideStruct.textBoxes = arrayNew(1) />
			<cfset slideStruct.images = arrayNew(1) />
			<cfset slideStruct.shapes = arrayNew(1) />
			<cfset slideStruct.lines = arrayNew(1) />
	
			<cfloop from="1" to="#arrayLen(thisSlideShapes)#" index="j">
				<cfset textBox = structNew() />
				<cfset line = structNew() />
				<cfset shape = structNew() />
				<cfset img = structNew() />
				<cfset thisSlideShape = thisSlideShapes[j] />
				<cfset anchor = thisSlideShape.getAnchor() />
				<cfset shapeClass = thisSlideShape.toString() />
				
				<cfif findNoCase("textbox", shapeClass)>
					<cfset textBox.textSpans = arrayNew(1)>
					<!--- if no text runs set to empty array --->
					<cftry>
						<cfset thisTextBoxTextSpans = fixNull(thisSlideShape.getTextRun().getRichTextRuns()) />
						<cfcatch>
							<cfset thisTextBoxTextSpans = arrayNew(1) />
						</cfcatch>
					</cftry>
					
					<cfloop from="1" to="#arrayLen(thisTextBoxTextSpans)#" index="j">
						
						<cfset textSpan = structNew() />
	
						<cfif thisTextBoxTextSpans[j].isItalic()>
							<cfset textSpan.fontStyle = "italic" />
						<cfelse>
							<cfset textSpan.fontStyle = "normal" />
						</cfif>
						<cfif thisTextBoxTextSpans[j].isBold()>
							<cfset textSpan.fontWeight = "bold" />
						<cfelse>
							<cfset textSpan.fontWeight = "normal" />
						</cfif>
						<cfif thisTextBoxTextSpans[j].isUnderlined()>
							<cfset textSpan.textDecoration = "underline" />
						<cfelse>
							<cfset textSpan.textDecoration = "none" />
						</cfif>
						
						<cfset textSpan.textAlign = getTextAlign(thisTextBoxTextSpans[j].getAlignment()) />
						
						<cfset textSpan.indentLevel = thisTextBoxTextSpans[j].getIndentLevel() />
						<!---
						i'm going to leave this out for now since it's goofy
						<cftry>
							<cfset textSpan.bulletChar = asc(thisTextBoxTextSpans[j].getBulletChar()) />
							<cfcatch>
								<cfset textSpan.bulletChar = "" />
							</cfcatch>
						</cftry>
						--->
						<cfset textSpan.fontSize = thisTextBoxTextSpans[j].getFontSize() />
						<cfset textSpan.fontFamily = thisTextBoxTextSpans[j].getFontName() />
						<cfset textSpan.fontColor = thisTextBoxTextSpans[j].getFontColor().getRed() & "," & thisTextBoxTextSpans[j].getFontColor().getGreen() & "," & thisTextBoxTextSpans[j].getFontColor().getBlue() />
						<cfset textSpan.text = htmlEditFormat(thisTextBoxTextSpans[j].getText()) /> 
						<cfset arrayAppend(textBox.textSpans, duplicate(textSpan)) />
					</cfloop>
					
					<cfset textBox.width = anchor.getWidth() />
					<cfset textBox.height = anchor.getHeight() />
					<cfset textBox.x = anchor.getX() />
					<cfset textBox.y = anchor.getY() />
					
					<cfset arrayAppend(slideStruct.textBoxes, duplicate(textBox)) />
				<cfelseif findNoCase("picture", shapeClass)>
					<cfset img.width = anchor.getWidth() />
					<cfset img.height = anchor.getHeight() />
					<cfset img.x = anchor.getX() />
					<cfset img.y = anchor.getY() />
					<cfset img.imgData = thisSlideShape.getPictureData().data />
					<cfset img.imgType = getImageType(thisSlideShape.getPictureData().type) />
					<cfset arrayAppend(slideStruct.images, duplicate(img)) />
				<cfelseif findNoCase("autoshape", shapeClass)>
					<cfset shape.width = anchor.getWidth() />
					<cfset shape.height = anchor.getHeight() />
					<cfset shape.x = anchor.getX() />
					<cfset shape.y = anchor.getY() />
					<cfset shape.shapeName = thisSlideShape.getShapeName() />
					<cfset arrayAppend(slideStruct.shapes, duplicate(shape)) />
				<cfelseif findNoCase("line", shapeClass)>		
					<cfset line.width = anchor.getWidth() />
					<cfset line.height = anchor.getHeight() />
					<cfset line.x = anchor.getX() />
					<cfset line.y = anchor.getY() />
					<cfset arrayAppend(slideStruct.lines, duplicate(line)) />
				</cfif>
			</cfloop>	
			
			<cfset arrayAppend(retArr, duplicate(slideStruct)) />		
		</cfloop>
		<cfreturn retArr />
	</cffunction>
	
	<cffunction name="getPPTMetaData" access="public" returntype="struct" output="false" hint="i extract metadata (author, saveinfo, etc) from a PPT">
		<cfargument name="pathToPPT" required="true" hint="the full path to the powerpoint to get the metadata for" />
		<cfset var hslf = instance.loader.create("org.apache.poi.hslf.HSLFSlideShow").init(arguments.pathToPPT) />
		<cfset var mdSummary = hslf.getSummaryInformation() />
		<cfset var md = structNew() />
		
		<cfset md.appName = fixNull(mdSummary.getApplicationName()) />
		<cfset md.author = fixNull(mdSummary.getAuthor()) />
		<cfset md.charCount = fixNull(mdSummary.getCharCount()) />
		<cfset md.comments = fixNull(mdSummary.getComments()) />
		<cfset md.createDateTime = fixNull(mdSummary.getCreateDateTime()) />
		<cfset md.editTime = fixNull(mdSummary.getEditTime()) />
		<cfset md.keywords = fixNull(mdSummary.getKeywords()) />
		<cfset md.lastAuthor = fixNull(mdSummary.getLastAuthor()) />
		<cfset md.lastPrinted = fixNull(mdSummary.getLastPrinted()) />
		<cfset md.lastSaveDateTime = fixNull(mdSummary.getLastSaveDateTime()) />
		<cfset md.pageCount = fixNull(mdSummary.getPageCount()) />
		<cfset md.revNumber = fixNull(mdSummary.getRevNumber()) />
		<cfset md.security = fixNull(mdSummary.getSecurity()) />
		<cfset md.subject = fixNull(mdSummary.getSubject()) />
		<cfset md.template = fixNull(mdSummary.getTemplate()) />
		<cfset md.thumbnail = fixNull(mdSummary.getThumbnail()) />
		<cfset md.title = fixNull(mdSummary.getTitle()) />
		<cfset md.wordCount = fixNull(mdSummary.getWordCount()) />
		
		<cfreturn md />
	</cffunction>
	<!--- private methods --->
	<cffunction name="fixNull" access="private" output="false" hint="internal private method to handle java nulls appropriately">
		<cfargument name="valueToFix" default="" />
		<cfset rStr = "" />
		<cfif isDefined("arguments.valueToFix")>
			<cfset rStr = arguments.valueToFix />
		</cfif>
		<cfreturn rStr />
	</cffunction>
	
	<cffunction name="getImageType" access="private" returntype="string" output="false" hint="i interpret the static constant pic type into a string">
		<cfargument name="picType" default="" />
		<cfset var type = "" />
		<cfswitch expression="#arguments.picType#">
			<cfcase value="2">
				<cfset type = "EMF" />
			</cfcase>
			<cfcase value="3">
				<cfset type = "WMF" />
			</cfcase>
			<cfcase value="4">
				<cfset type = "PICT" />
			</cfcase>
			<cfcase value="5">
				<cfset type = "JPEG" />
			</cfcase>
			<cfcase value="6">
				<cfset type = "PNG" />
			</cfcase>
			<cfcase value="7">
				<cfset type = "DIB" />
			</cfcase>
			<cfdefaultcase>
				<cfset type = "" />
			</cfdefaultcase>
		</cfswitch>
		<cfreturn type />
	</cffunction>
	
	<cffunction name="getTextAlign" access="private" output="false" hint="i interpret the text alignment from the static type into a descriptive string">
		<cfargument name="textAlign" default="" hint="the static value returned by java" />
		<cfset var align = "" />
		<cfswitch expression="##">	
			<cfcase value="0">
				<cfset align = "left" />
			</cfcase>
			<cfcase value="1">
				<cfset align = "center" />
			</cfcase>
			<cfcase value="2">
				<cfset align = "right" />
			</cfcase>
			<cfcase value="3">
				<cfset align = "justify" />
			</cfcase>
			<cfdefaultcase>
				<cfset align = "left" />
			</cfdefaultcase>
		</cfswitch>
		<cfreturn align />
	</cffunction>
</cfcomponent>
