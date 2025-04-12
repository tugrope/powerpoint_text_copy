-- ���̃X�N���v�g�́AMicrosoft PowerPoint�@for mac �̃A�N�e�B�u�v���[���e�[�V��������e�L�X�g�{�b�N�X�̃e�L�X�g�����W���A�N���b�v�{�[�h�ɕۑ����܂��B
-- �O���[�v�����ꂽ�V�F�C�v�̓R�s�[�ł��Ȃ��̂ŁA�O���[�v�����ꂽ�V�F�C�v�̓O���[�v�������Ă���R�s�[���܂��B
set the clipboard to ""

tell application "Microsoft PowerPoint"
	set collectedText to ""

	-- �A�N�e�B�u�v���[���e�[�V�������擾���܂��B
	set activePresentation to active presentation

	-- �e�X���C�h�����[�v�������܂��B
	repeat with slideIndex from 1 to count slides of activePresentation
		set currentSlide to slide slideIndex of activePresentation

		-- �e�V�F�C�v�����[�v�������܂��B
		repeat with shapeIndex from 1 to count shapes of currentSlide
			set currentShape to shape shapeIndex of currentSlide

			-- �e�L�X�g�t���[�������݂��邩�m�F���܂��B
			if has text frame of currentShape is true then
				set textContent to text range of text frame of currentShape
				if textContent is not missing value then
					set collectedText to collectedText & (content of textContent) & linefeed
				end if
			end if
		end repeat
	end repeat
end tell

-- ���W�����e�L�X�g���N���b�v�{�[�h�ɐݒ肵�܂��B
set the clipboard to collectedText
display dialog "�e�L�X�g�{�b�N�X�̃e�L�X�g���N���b�v�{�[�h�ɕۑ�����܂����B"
