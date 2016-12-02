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
      attr_reader :theme_overrides
      attr_reader :notes
      attr_reader :tags
      attr_reader :drawings
      attr_reader :notes_slides
      attr_reader :master
      attr_reader :notes_master
      attr_reader :layouts
      attr_reader :layout
      attr_reader :file_types

      def initialize(options={})
        require_arguments [:presentation, :title, :content, :rel_content, :images, :charts, :embeddings, :notes, :tags, :drawings, :master, :notes_master, :layout, :theme_overrides], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @file_types = []
        @notes_slides = []
        puts theme_overrides

      end

      def save(extract_path, index)
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
        save_images(extract_path, index) if images && images.count > 0
        save_theme_overrides(extract_path, index) if theme_overrides && theme_overrides.count > 0
        save_charts(extract_path, index) if charts && charts.count > 0
        save_embeddings(extract_path, index) if embeddings && embeddings.count > 0
        save_notes(extract_path, index) if notes && notes.count > 0
        save_tags(extract_path, index) if tags && tags.count > 0
        save_drawings(extract_path, index) if drawings && drawings.count > 0
      end

      def file_type
        file_types
      end

      private

      def save_rel_xml(extract_path, index)
        @index = index
        @tmp_content = rel_content.to_s
        @tmp_content.gsub!('charts',"charts/slide_#{@index}")
        @tmp_content.gsub!('media',"media/slide_#{@index}")
        @tmp_content.gsub!('../tags',"../tags/slide_#{@index}")
        xml = Nokogiri::XML::Document.parse @tmp_content
        xml.css('Relationship').select{ |node|
          if node['Type'].include? 'relationships/notesSlide'
            node['Target'] = "../notesSlides/notesSlide#{index}.xml"
          elsif node['Type'].include? 'relationships/masterSlide'
            node['Target'] = master[:file_path]
          elsif node['Type'].include? 'relationships/slideLayout'
            node['Target'] = layout[:file_path]
          end
        }
        @tmp_content = xml.to_s
        render_view('ff_embeded_slide_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
      end

      def save_slide_xml(extract_path, index)
        @index = index
        render_view('ff_embeded_slide_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end

      def save_images(extract_path, index)
        images.each do |image|
          FileUtils::mkdir_p "#{extract_path}/ppt/media/slide_#{index}"
          zip_entry = image.rewind
          file_path = zip_entry.name.to_s.gsub('media',"media/slide_#{index}")
          File.open("#{extract_path}/" + file_path, 'wb') do |f|
            f.write zip_entry.get_input_stream.read
          end
          @file_types << { type: MimeMagic.by_magic(File.open("#{extract_path}/" + file_path)).type, path: "/#{file_path}" } unless file_path.include? "rels"
          image.close
        end
      end

      def save_charts(extract_path, index)
        FileUtils::mkdir_p "#{extract_path}/ppt/charts/slide_#{index}/_rels"
        charts.each do |chart|
          zip_entry = chart.rewind
          file_path = zip_entry.name.to_s.gsub('charts',"charts/slide_#{index}")
          File.open("#{extract_path}/" +  file_path, 'wb') do |f|
            f.write zip_entry.get_input_stream.read
              .gsub('../embeddings',"../../embeddings/slide_#{index}")
              .gsub('../drawings',"../../drawings/slide_#{index}")
              .gsub('../theme',"../../theme/slide_#{index}")
          end
          @file_types << { type: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml" , path: "/#{file_path}" } unless file_path.include? "rels"
          chart.close
        end
      end

      def save_embeddings(extract_path, index)
        FileUtils::mkdir_p "#{extract_path}/ppt/embeddings/slide_#{index}"
        embeddings.each do |embedding|
          zip_entry = embedding.rewind
          File.open("#{extract_path}/" + zip_entry.name.to_s.gsub('embeddings',"embeddings/slide_#{index}"), 'wb') do |f|
            f.write zip_entry.get_input_stream.read
          end
          embedding.close
        end
      end

      def save_drawings(extract_path, index)
        FileUtils::mkdir_p "#{extract_path}/ppt/drawings/slide_#{index}"
        drawings.each do |drawing|
          zip_entry = drawing.rewind
          file_path = zip_entry.name.to_s.gsub('drawings/',"drawings/slide_#{index}/")
          File.open("#{extract_path}/" + file_path, 'wb') do |f|
            f.write zip_entry.get_input_stream.read
          end
          @file_types << { type: "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml", path: "/#{file_path}" } unless file_path.include? "rels"
          drawing.close
        end
      end

      def save_theme_overrides(extract_path, index)
        FileUtils::mkdir_p "#{extract_path}/ppt/theme/slide_#{index}"
        theme_overrides.each do |theme_override|
          zip_entry = theme_override.rewind
          puts zip_entry
          file_path = zip_entry.name.to_s.gsub('ppt/theme/',"ppt/theme/slide_#{index}/")

          puts file_path
          File.open("#{extract_path}/" + file_path, 'wb') do |f|
            f.write zip_entry.get_input_stream.read
          end
          @file_types << { type: "application/vnd.openxmlformats-officedocument.themeOverride+xml", path: "/#{file_path}" } unless file_path.include? "rels"
          theme_override.close
        end
      end

      def save_tags(extract_path, index)
        FileUtils::mkdir_p "#{extract_path}/ppt/tags/slide_#{index}"
        tags.each do |tag|
          zip_entry = tag.rewind
          file_path = zip_entry.name.to_s.gsub('ppt/tags/',"ppt/tags/slide_#{index}/")
          File.open("#{extract_path}/" + file_path, 'wb') do |f|
            f.write zip_entry.get_input_stream.read
          end
          @file_types << { type: "application/vnd.openxmlformats-officedocument.presentationml.tags+xml", path: "/#{file_path}" } unless file_path.include? "rels"
          tag.close
        end
      end

      def save_notes(extract_path, index)
        notes.each do |note|
          zip_entry = note.rewind
          if zip_entry.name.include? "rels"
            notes_xml = Nokogiri::XML::Document.parse zip_entry.get_input_stream.read
            notes_xml.css('Relationship').select{ |node|
              if node['Type'].include? 'relationships/slide'
                node['Target'] = "../slides/slide#{index}.xml"
              elsif node['Type'].include? 'relationships/notesMaster'
                node['Target'] = notes_master[:file_path]
              end
            }
            File.open("#{extract_path}/ppt/notesSlides/_rels/notesSlide#{index}.xml.rels", 'wb') do |f|
              f.write notes_xml
            end
          else
            file_path = "ppt/notesSlides/notesSlide#{index}.xml"
            @notes_slides << { type: "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml", path: "#{file_path}" }
            notes_xml =  Nokogiri::XML::Document.parse  zip_entry.get_input_stream.read
            File.open("#{extract_path}/" + file_path , 'wb') do |f|
              f.write notes_xml.to_xml.gsub('smtClean="0"','')
            end
          end
          note.close
        end
      end
    end
  end
end
