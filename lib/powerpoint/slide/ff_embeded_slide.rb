require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'
require 'mimemagic'

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
      attr_reader :notes
      attr_reader :notes_slides
      attr_reader :file_types

      def initialize(options={})
        require_arguments [:presentation, :title, :content, :rel_content, :images, :charts, :embeddings, :notes], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @file_types = []
        @notes_slides = []
      end

      def save(extract_path, index)
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
        save_images(extract_path, index) if images && images.count > 0
        save_charts(extract_path, index) if charts && charts.count > 0
        save_embeddings(extract_path, index) if embeddings && embeddings.count > 0
        save_notes(extract_path, index) if notes && notes.count > 0
      end

      def file_type
        file_types
      end

      def save_rel_xml(extract_path, index)
        @index = index
        @tmp_content = rel_content.to_s
        @tmp_content.gsub!('slideLayouts','slideLayouts/master2')
        @tmp_content.sub!('charts',"charts/slide_#{@index}")
        @tmp_content.gsub!('media',"media/slide_#{@index}")
        @tmp_content.gsub!('embeddings',"media/slide_#{@index}")
        xml = Nokogiri::XML::Document.parse @tmp_content
        xml.css('Relationship').select{ |node|
          if node['Type'].include? 'relationships/notesSlide'
            node['Target'] = "../notesSlides/notesSlide#{index}.xml"
          end
        }
        @tmp_content = xml.to_s
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
          file_path = zip_entry.name.to_s.gsub('media',"media/slide_#{index}")
          File.open("#{extract_path}/" + file_path, 'wb') do |f|
            f.write zip_entry.get_input_stream.read
          end
          f.close
          @file_types << { type: MimeMagic.by_magic(File.open("#{extract_path}/" + file_path)).type, path: "/#{file_path}" } unless file_path.include? "rels"
        end
      end
      private :save_images

      def save_charts(extract_path, index)
        FileUtils::mkdir_p "#{extract_path}/ppt/charts/slide_#{index}/_rels"
        charts.each do |chart|
          zip_entry = chart.rewind
          file_path =zip_entry.name.to_s.gsub('charts',"charts/slide_#{index}")
          File.open("#{extract_path}/" +  file_path, 'wb') do |f|
            f.write zip_entry.get_input_stream.read.gsub('../embeddings',"../../embeddings/slide_#{index}")
          end
          f.close
          @file_types << { type: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml" , path: "/#{file_path}" } unless file_path.include? "rels"
        end
      end
      private :save_charts

      def save_embeddings(extract_path, index)
        FileUtils::mkdir_p "#{extract_path}/ppt/embeddings/slide_#{index}"
        embeddings.each do |embedding|
          zip_entry = embedding.rewind
          File.open("#{extract_path}/" + zip_entry.name.to_s.gsub('embeddings',"embeddings/slide_#{index}"), 'wb') do |f|
            f.write zip_entry.get_input_stream.read
          end
          f.close
        end
      end
      private :save_embeddings

      def save_notes(extract_path, index)
        notes.each do |note|
          zip_entry = note.rewind
          if zip_entry.name.include? "rels"
            notes_xml = Nokogiri::XML::Document.parse zip_entry.get_input_stream.read
            notes_xml.css('Relationship').select{ |node|
              if node['Type'].include? 'relationships/slide'
                node['Target'] = "../slides/slide#{index}.xml"
              end
            }
            File.open("#{extract_path}/ppt/notesSlides/_rels/notesSlide#{index}.xml.rels", 'wb') do |f|
              f.write notes_xml
            end
            f.close
          else
            notes_file = "ppt/notesSlides/notesSlide#{index}.xml"
            notes_slides << "/#{notes_file}"
            File.open("#{extract_path}/" + notes_file , 'wb') do |f|
              f.write zip_entry.get_input_stream.read
            end
            f.close
          end
        end
      end
      private :save_notes

    end
  end
end
