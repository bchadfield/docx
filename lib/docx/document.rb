require 'docx/parser'
require 'zip/zip'
require 'html_writer'

module Docx
  # The Document class wraps around a docx file and provides methods to
  # interface with it.
  # 
  #   # get a Docx::Document for a docx file in the local directory
  #   doc = Docx::Document.open("test.docx")
  #   
  #   # get the text from the document
  #   puts doc.text
  #   
  #   # do the same thing in a block
  #   Docx::Document.open("test.docx") do |d|
  #     puts d.text
  #   end
  class Document
    delegate :paragraphs, :bookmarks, :tables, :to => :@parser
    delegate :doc, :xml, :zip, :to => :@parser
    def initialize(path, &block)
      @replace = {}
      if block_given?
        @parser = Parser.new(File.expand_path(path), &block)
      else
        @parser = Parser.new(File.expand_path(path))
      end
    end
    
    # With no associated block, Docx::Document.open is a synonym for Docx::Document.new. If the optional code block is given, it will be passed the opened +docx+ file as an argument and the Docx::Document oject will automatically be closed when the block terminates. The values of the block will be returned from Docx::Document.open.
    # call-seq:
    #   open(filepath) => file
    #   open(filepath) {|file| block } => obj
    def self.open(path, &block)
      self.new(path, &block)
    end

    ##
    # *Deprecated*
    # 
    # Iterates over paragraphs within document
    # call-seq:
    #   each_paragraph => Enumerator
    def each_paragraph
      paragraphs.each { |p| yield(p) }
    end
    
    # call-seq:
    #   to_s -> string
    def to_s
      paragraphs.map(&:to_s).join("\n")
    end

    # Save document to provided path
    # call-seq:
    #   save(filepath) => void
    def save(path)
      update
      Zip::ZipOutputStream.open(path) do |out|
        zip.each do |entry|
          out.put_next_entry(entry.name)

          if @replace[entry.name]
            out.write(@replace[entry.name])
          else
            out.write(zip.read(entry.name))
          end
        end
      end
      zip.close
    end

    def to_html(options={})
      HtmlWriter.new.write do |html|
        html.doctype options[:doctype] || 5
        html.head do |head|
          head.title options[:title] || ''
        end
        html.body do |body|
          self.each_paragraph do |paragraph|
            if !paragraph.text_runs.empty? && paragraph.text_runs[0].heading?
              body.send("h#{paragraph.text_runs[0].heading[-1]}", paragraph.text_runs[0].text)
            else
              body.p paragraph.text_runs.map{ |tr| inline_content_for(tr) }.join('')
            end
          end
        end
      end
    end
    
    alias_method :text, :to_s

    private

    #--
    # TODO: Flesh this out to be compatible with other files
    # TODO: Method to set flag on files that have been edited, probably by inserting something at the 
    # end of methods that make edits?
    #++
    def update
      @replace["word/document.xml"] = doc.serialize :save_with => 0
    end

    def inline_content_for(text_run)
      inline_content = text_run.text
      inline_content = wrap(inline_content, :em)     if text_run.italicized?
      inline_content = wrap(inline_content, :strong) if text_run.bolded?
      inline_content = wrap(inline_content, :span,
        {style: 'text-decoration: underline;'})      if text_run.underlined?
      inline_content
    end
    
    def wrap(text, tag, attrs=nil)
      html = "<#{tag.to_s}"
      unless attrs.nil?
        attrs.each { |k,v| html << " #{k.to_s}=\"#{v.to_s}\""}
      end
      html << ">#{text}</#{tag.to_s}>"
    end

  end
end
