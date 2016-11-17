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
      media = Array.new
      return media unless @files
      @files.each do |f|
        if f.name.include? 'ppt/media'
          media.push(f.name)
        end
      end
      media
    end

    def charts
      charts = Array.new
      return charts unless @files
      @files.each do |f|
        if f.name.include? 'ppt/charts'
          charts.push(f.name)
        end
      end
      charts
    end

    def embeddings
      embeddings = Array.new
      return embeddings unless @files
      @files.each do |f|
        if f.name.include? 'ppt/embeddings'
          embeddings.push(f.name)
        end
      end
      embeddings
    end

    def close
      @files.close
    end
  end
end
