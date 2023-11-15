require 'google_drive'

class RadniList
      include Enumerable
      attr_accessor :worksheet
      attr_accessor :red_to_ignore

      def initialize(worksheet_index)
            session = GoogleDrive::Session.from_config('config.json')
            spreadsheet = session.spreadsheet_by_key('1prp_g537u6729MQ3gOHqlIO3YayRQakhDDiZ7VqW5LM')
            @worksheet = spreadsheet.worksheets[worksheet_index]
            @red_to_ignore = -1
            check
            add_kolone_metode
      end

      def row(index)
            if index != @red_to_ignore
                  @worksheet.rows[index - 1]
            end
      end

      def get_dvo_dim_niz
            @worksheet.rows
      end

      def each
            get_dvo_dim_niz.each do |red|
                  red.each do |celija|
                        yield celija
                  end
            end
      end

      def [](kolona_ime)
            # kolona_index = worksheet.rows[0].index(kolona_ime) + 1
            KolonaProxy.new(self, kolona_ime)

            # put_col(kolona_index)
      end

      def put_col(kolona_index)
            col_to_put = []
            i = 0
            
            while i < worksheet.rows.size
                  col_to_put << worksheet[i + 1, kolona_index + 1]
                  i += 1
            end

            col_to_put
      end

      def check
            get_dvo_dim_niz.each_with_index do |red, red_index|
                  red.each_with_index do | celija, kolona_index|
                        if celija == "total" or 
                           celija == "subtotal"
                              @red_to_ignore = red_index + 1
                        end
                  end
            end
      end

      def -(radni_list2)
            brojac = 2
            if radni_list2.is_a?(RadniList)
                  do_step_minus(brojac, radni_list2)
            else
                  raise ArgumentError, "Greska ne mogu se oduzimati #{radni_list2.class} objekti"
            end
      end

      def do_step_minus(brojac, radni_list2)
            get_dvo_dim_niz.each_with_index do |red, index|
                  next if index.zero?
                  next if index == 1
                  radni_list2.get_dvo_dim_niz.each do |red2|
                        if red == red2
                              self.worksheet.delete_rows(brojac, 1)
                              brojac -= 1
                        end
                  end
                  brojac += 1
            end
                  self.worksheet.save
      end
        
      def +(radni_list2)
            list = []
            n = 0
            if radni_list2.is_a?(RadniList)
                  do_step_plus(list, n, radni_list2)
            else
                  raise ArgumentError, "Greska ne mogu se oduzimati #{radni_list2.class} objekti"
            end
      end

      def do_step_plus(
            list,
            n,
            radni_list2)
            get_dvo_dim_niz.each_with_index do |red,index|
                  next if index.zero?
                  next if index == 1
                  list << red
                  n += 1
            end
            radni_list2.get_dvo_dim_niz.each_with_index do |red2,index2|
                  next if index2.zero?
                  next if index2 == 1
                  list << red2 unless list.include?(red2)
            end
            list.shift(n)
      
            self.worksheet.insert_rows(3, list)
      
            self.worksheet.save
      end

      private

      def add_kolone_metode
            prvi_red = @worksheet.rows[0]
            prvi_red.each do |kolona_ime|
                  instance_eval do
                        define_singleton_method(normalize_method_name(kolona_ime)) do
                              KolonaProxy.new(self, kolona_ime)
                        end
                  end
            end
      end
        
      def normalize_method_name(name)
            name.downcase.split.map(&:capitalize).join('')
      end

      class KolonaProxy
            def initialize(
                  radni_list,
                  kolona_ime
                  )
                  @radni_list = radni_list
                  @kolona_ime = kolona_ime
                  @kolona_index = @radni_list.worksheet.rows[0].index(@kolona_ime) + 1
                  add_celija_metode
            end

            def add_celija_metode
                  celija = []
                  @radni_list.worksheet.rows.each do |red|
                        celija << red[@kolona_index - 1]
                  end

                  row_to_give = nil

                  celija.each_with_index do |ime_celija, celija_index|
                        instance_eval do
                              define_singleton_method(ime_celija.downcase) do
                                    row_to_give = @radni_list.row(celija_index + 1)
                              end
                        end
                  end

                  if row_to_give != nil
                        row_to_give
                  end
            end

            def [](red_index) 
                  if red_index != @radni_list.red_to_ignore
                        do_step(red_index)
                  end

            end

            def do_step(red_index)
                  kolona_index = @radni_list.worksheet.rows[0].index(@kolona_ime)
                  @radni_list.worksheet[red_index, kolona_index + 1]
            end

            def []=(
                  red_index,
                  nova_vrednost
                  )
                  if red_index != @radni_list.red_to_ignore
                        do_step2
                  end
            end

            def do_step2
                  kolona_index = @radni_list.worksheet.rows[0].index(@kolona_ime)
                  @radni_list.worksheet[red_index, kolona_index + 1] = nova_vrednost
                  @radni_list.worksheet.save
            end

            def sum
                  sum = 0
                  @radni_list.worksheet.rows.each do |red|
                        celija = red[@kolona_index - 1]
                        sum += celija.to_i if celija.to_i.to_s == celija
                  end
                  sum
            end

            def avg
                  sum = 0.0
                  num = 0
                  @radni_list.worksheet.rows.each do |red|
                        element = red[@kolona_index - 1]
                        broj = element.to_i
                        if broj.to_s == element
                              sum += broj
                              num += 1
                        end
                  end

                  num.zero? ? 0 : sum / num
            end
      end
end

indeks_lista = 0

radni_list = RadniList.new(indeks_lista)
radni_list2 = RadniList.new(indeks_lista + 1)

# p radni_list.get_dvo_dim_niz # 1.

# p radni_list.row(1) # 2.

# radni_list.get_dvo_dim_niz.each do |e|
#       p e                                # 3.
# end

# p radni_list["Prva Kolona"] # 5. a

# p radni_list["Prva Kolona"][3] # 5. b

# radni_list["Prva Kolona"][3] = 2556 # 5. c

# p radni_list.PrvaKolona # 6. a

# p radni_list.PrvaKolona.sum # 6. i
# p radni_list2.DrugaKolona.avg # 6. i

# p radni_list.TrecaKolona.aa # 6. ii

# p radni_list.red_to_ignore # 7.

# radni_list + radni_list2 # 8.

# radni_list - radni_list2 # 9. 

# 10.