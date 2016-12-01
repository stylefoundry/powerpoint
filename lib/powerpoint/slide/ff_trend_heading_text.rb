require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'
require 'mimemagic'

module Powerpoint
    module Slide
    class FFTrendHeadingText
      include Powerpoint::Util

      attr_reader :title, :content, :image_paths, :images

      def initialize(options={})
        require_arguments [:presentation, :title, :content, :image_paths], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @images = image_paths.each.map { |image_path|  [ File.basename(image_path), image_path ] }
      end

      def file_type
        @images.map{ |image_name, image_path| { type: MimeMagic.by_magic(File.open(image_path)).type, path: "/ppt/media/#{image_name}" } }
      end

      def save(extract_path, index)
        @images.each do |image_path, image_name|
          copy_media(extract_path, image_path) if image_path != nil
        end
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def save_rel_xml(extract_path, index)
        render_view('ff_trend_heading_text_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('ff_trend_heading_text_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml

    end
  end
end
