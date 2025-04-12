-- このスクリプトは、Microsoft PowerPoint　for mac のアクティブプレゼンテーションからテキストボックスのテキストを収集し、クリップボードに保存します。
-- グループ化されたシェイプはコピーできないので、グループ化されたシェイプはグループ解除してからコピーします。
set the clipboard to ""

tell application "Microsoft PowerPoint"
	set collectedText to ""

	-- アクティブプレゼンテーションを取得します。
	set activePresentation to active presentation

	-- 各スライドをループ処理します。
	repeat with slideIndex from 1 to count slides of activePresentation
		set currentSlide to slide slideIndex of activePresentation

		-- 各シェイプをループ処理します。
		repeat with shapeIndex from 1 to count shapes of currentSlide
			set currentShape to shape shapeIndex of currentSlide

			-- テキストフレームが存在するか確認します。
			if has text frame of currentShape is true then
				set textContent to text range of text frame of currentShape
				if textContent is not missing value then
					set collectedText to collectedText & (content of textContent) & linefeed
				end if
			end if
		end repeat
	end repeat
end tell

-- 収集したテキストをクリップボードに設定します。
set the clipboard to collectedText
display dialog "テキストボックスのテキストがクリップボードに保存されました。"
