require "docx_replace/version"
require 'zip'
require 'tempfile'

module DocxReplace
  class Doc
    DOCUMENT_FILE_PATH = 'word/document.xml'

    def initialize(source, target, &block)
      @zip_file = Zip::File.new(source)
      @temp_file = Tempfile.new('docxedit-')
      @document_content = @zip_file.read(DOCUMENT_FILE_PATH)
      @target = target
      block.call(self) if block_given?
    end

    def replace(pattern, replacement, multiple_occurrences=false)
      if multiple_occurrences
        @document_content.gsub!(pattern, replacement)
      else
        @document_content.sub!(pattern, replacement)
      end
    end

    def commit
      Zip::OutputStream.open(@temp_file.path) do |zos|
        @zip_file.entries.each do |e|
          unless e.name == DOCUMENT_FILE_PATH
            zos.put_next_entry(e.name)
            zos.print e.get_input_stream.read
          end
        end

        zos.put_next_entry(DOCUMENT_FILE_PATH)
        zos.print @document_content
      end

      FileUtils.mv(@temp_file.path, @target)
      @temp_file.close
      @temp_file.unlink
      @target
    end
  end
end
