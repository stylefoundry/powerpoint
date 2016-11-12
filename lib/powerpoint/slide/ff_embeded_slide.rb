require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'

module Powerpoint
  module Slide
    class FFEmbededSlide
      include Powerpoint::Util

      attr_reader :content
      attr_reader :rel_content
      attr_reader :title
      attr_reader :images
      attr_reader :charts
      attr_reader :embeddings

      def initialize(options={})
        require_arguments [:presentation, :title, :content, :rel_content, :images, :charts, :embeddings], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
      end

      def save(extract_path, index)
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
        save_images(extract_path, index)
        save_charts(extract_path, index)
        save_embeddings(extract_path, index)
      end

      def file_type
        nil
      end

      def save_rel_xml(extract_path, index)
        @index = index
        @tmp_content = rel_content.to_s
        @tmp_content.gsub!('slideLayouts','slideLayouts/master2')
        @tmp_content.sub!('charts',"charts/slide_#{@index}")
        @tmp_content.gsub!('media',"media/slide_#{@index}")
        @tmp_content.gsub!('embeddings',"media/slide_#{@index}")
        render_view('ff_embeded_slide_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        @index = index
        render_view('ff_embeded_slide_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml

      def save_images(extract_path, index)
        images.each do |image|
          FileUtils::mkdir_p "#{extract_path}/ppt/media/slide_#{index}"
          zip_entry = image.rewind
          File.open("#{extract_path}/" + zip_entry.name.to_s.gsub('media',"media/slide_#{index}"), 'w+') do |f|
            f.write zip_entry.get_input_stream.read
          end
        end
      end
      private :save_images

      def save_charts(extract_path, index)
        FileUtils::mkdir_p "#{extract_path}/ppt/charts/slide_#{index}/_rels"
        charts.each do |chart|
          zip_entry = chart.rewind
          File.open("#{extract_path}/" + zip_entry.name.to_s.gsub('charts',"charts/slide_#{index}/"), 'w+') do |f|
            f.write zip_entry.get_input_stream.read.gsub('../embeddings',"../../embeddings/slide_#{index}")
          end
        end
      end
      private :save_charts

      def save_embeddings(extract_path, index)
        FileUtils::mkdir_p "#{extract_path}/ppt/embeddings/slide_#{index}"
        embeddings.each do |embedding|
          zip_entry = embedding.rewind
          puts zip_entry.name
          File.open("#{extract_path}/" + zip_entry.name.to_s.gsub('embeddings',"embeddings/slide_#{index}/"), 'w+') do |f|
            f.write zip_entry.get_input_stream.read
          end
        end
      end
      private :save_embeddings

    end
  end
end
