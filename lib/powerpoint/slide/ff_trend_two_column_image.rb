require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'
require 'mimemagic'

module Powerpoint
    module Slide
    class FFTrendTwoColumnImage
      include Powerpoint::Util

      attr_reader :title, :content, :question, :images, :link_path

      def initialize(options={})
        require_arguments [:presentation, :title, :content, :question, :images, :link_path], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @images = images.each.map { |image_path|  [ File.basename(image_path), image_path ] }
      end

      def file_type
        @images.map{ |image_name, image_path| { type: MimeMagic.by_magic(File.open(image_path)).type, path: "/ppt/media/#{image_name.gsub('jpg','jpeg')}" } }
      end

      def save(extract_path, index)
        @images.each do |image_name, image_path|
          copy_media(extract_path, image_path) if image_path != nil
        end        
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def save_rel_xml(extract_path, index)
        render_view('ff_trend_two_column_image_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('ff_trend_two_column_image_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml

    end
  end
end
