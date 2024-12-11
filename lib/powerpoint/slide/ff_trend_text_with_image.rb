require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'
require 'mimemagic'
require 'nokogiri'

module Powerpoint
    module Slide
    class FFTrendTextWithImage
      include Powerpoint::Util

      attr_reader :title, :content, :image_path, :image, :links, :source

      def initialize(options={})
        require_arguments [:presentation], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}

        @image = [File.basename(image_path), image_path] if image_path
      end

      def file_type
        if image
          image_name, image_path = image
          [
            {
              type: MimeMagic.by_magic(File.open(image_path)).type,
              path: "/ppt/media/#{image_name}"
            }
          ]
        else
          []
        end
      end

      def save(extract_path, index)
        if image
          copy_media(extract_path, image[1])
        end
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def save_rel_xml(extract_path, index)
        render_view('ff_text_with_image_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('ff_text_with_image_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml

    end
  end
end
