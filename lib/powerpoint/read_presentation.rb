require 'zip/filesystem'
require 'nokogiri'

module Powerpoint

  class Powerpoint::ReadPresentation

    attr_reader :files

    def initialize path
      raise 'Not a valid file format.' unless (['.pptx'].include? File.extname(path).downcase)
      @files = Zip::File.open path
    end

    def slides
      slides = Array.new
      return slides unless @files
      @files.each do |f|
        if f.name.include? 'ppt/slides/slide'
          slides.push Powerpoint::ReadSlide.new(self, f.name)
        end
      end
      slides.sort{|a,b| a.slide_num <=> b.slide_num}
    end
  
    def media
      match_files 'ppt/media'
    end

    def charts
      match_files 'ppt/charts'
    end

    def embeddings
      match_files 'ppt/embeddings'
    end

    def masters
      match_files('ppt/slideMasters').select{ |file_name| file_name unless file_name.include? 'rels' }
    end

    def notes_masters
      match_files('ppt/notesMasters').select{ |file_name| file_name unless file_name.include? 'rels' }
    end

    def layouts
      match_files('ppt/slideLayouts').select{ |file_name| file_name unless file_name.include? 'rels' }
    end

    def close
      @files.close
    end

    private

    def match_files(pattern)
      matches = Array.new
      return matches unless @files
      @files.each do |f|
        if f.name.include? pattern
          matches.push(f.name)
        end
      end
      matches
    end
  end
end
