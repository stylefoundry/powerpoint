require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'
require 'mimemagic'

module Powerpoint
  module Slide
    class FFTrendIntro
      include Powerpoint::Util

      attr_reader :image_name, :title, :subtitle, :image_path, :trend_number

      def initialize(options={})
        require_arguments [:presentation, :title, :subtitle, :image_path], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        resize_image!
        @image_name = File.basename(@image_path) if @image_path != nil && @image_path != ""
      end

      def save(extract_path, index)
        copy_media(extract_path, @image_path) if @image_path != nil && @image_name != ""
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def file_type
        [{ type: MimeMagic.by_magic(File.open(image_path)).type, path: "/ppt/media/#{image_name}" }]
      end

      def resize_image!
        unless image_path && File.file?(image_path)
          return nil
        end

        # See presentation.xml.erb p:sldSz cx/cy
        slide_width = pt_to_pixle 12192000
        slide_height = pt_to_pixle 6858000

        image = Magick::ImageList.new(image_path).first
        image_height = image.rows
        image_width = image.columns
        image_ratio = image_width / image_height.to_f

        width = slide_width
        height = slide_width / image_ratio

        if height < slide_height
          height = slide_height
          width = slide_height * image_ratio
        end

        image.resize!(width, height)
        image.crop!(Magick::CenterGravity, slide_width, slide_height)
        image.write(image_path)
      end
      private :resize_image!

      def save_rel_xml(extract_path, index)
        render_view('ff_trend_intro_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('ff_trend_intro_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end
end
