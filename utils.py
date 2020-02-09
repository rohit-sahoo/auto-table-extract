from __future__ import division
from pdfminer.layout import LAParams
import pandas as pd
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.layout import LAParams
from pdfminer.converter import PDFPageAggregator
import pdfminer
from pdfminer.pdfdevice import PDFDevice
from sklearn.cluster import KMeans
from os.path import basename
import re
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import resolve1
import os
import csv



def helper_anomaly(param):
    pass


def auto_table_extract(example_file):

    all_tables = list()
    #example_file = r"C:\Users\divesh.kubal\Downloads\H& B PDF\H& B PDF\16154.pdf"
    #my_pass = '1234'
    file = open(example_file, 'rb')
    parser = PDFParser(file)
    #document = PDFDocument(parser,password=my_pass)
    document = PDFDocument(parser)
    total_pages = resolve1(document.catalog['Pages'])['Count']
    # print('page numbers: ', total_pages)

    total_pages = resolve1(document.catalog['Pages'])['Count']
    base_filename = basename(example_file)
    bs= base_filename
    #page_number1 = int(input('Enter Page Number: '))
    #page_number = page_number1 - 1
    #base_filename = base_filename.replace('.pdf','') + '_pg_' + str(page_number1)
    f = open('math_log.txt', 'a', encoding='utf-8')

    number_of_clusters_list = []


    for page_number in range(0,total_pages):


        base_filename = base_filename.replace('.pdf', '') + '_pg_' + str(page_number)
        class pdfPositionHandling:
            xo = list()

            yo = list()
            text = list()
            def parse_obj(self, lt_objs):

                # loop over the object list
                for obj in lt_objs:

                    if isinstance(obj, pdfminer.layout.LTTextLine):
                        pdfPositionHandling.xo.append(int(obj.bbox[0]))
                        pdfPositionHandling.yo.append(int(obj.bbox[1]))
                        pdfPositionHandling.text.append(str(obj.get_text()))


                        math_log = str(obj.bbox[0]) + ' ' + str(obj.bbox[1]) + ' ' + str(obj.get_text().replace('\n', '_'))
                        f.write(math_log + '\n')
                    # if it's a textbox, also recurse

                    if isinstance(obj, pdfminer.layout.LTTextBoxHorizontal):
                        self.parse_obj(obj._objs)

                    # if it's a container, recurse
                    elif isinstance(obj, pdfminer.layout.LTFigure):
                        self.parse_obj(obj._objs)



            def parsepdf(self, filename, startpage, endpage):

                # Open a PDF file.
                fp = open(filename, 'rb')

                # Create a PDF parser object associated with the file object.
                parser = PDFParser(fp)

                # Create a PDF document object that stores the document structure.
                # Password for initialization as 2nd parameter
                document = PDFDocument(parser)


                # Check if the document allows text extraction. If not, abort.
                if not document.is_extractable:
                    raise PDFTextExtractionNotAllowed

                # Create a PDF resource manager object that stores shared resources.
                rsrcmgr = PDFResourceManager()

                # Create a PDF device object.
                device = PDFDevice(rsrcmgr)

                # BEGIN LAYOUT ANALYSIS
                # Set parameters for analysis.
                laparams = LAParams()

                # Create a PDF page aggregator object.
                device = PDFPageAggregator(rsrcmgr, laparams=laparams)

                    # Create a PDF interpreter object.
                interpreter = PDFPageInterpreter(rsrcmgr, device)


                i = 0
                # loop over all pages in the document
                for page in PDFPage.create_pages(document):
                    if i >= startpage and i <= endpage:
                        # read the page into a layout object
                        interpreter.process_page(page)
                        layout = device.get_result()

                        # extract text from this object
                        self.parse_obj(layout._objs)
                    i += 1


        def table_without_border():

            obj = pdfPositionHandling()
            obj.parsepdf(r'input_pdf.pdf', 0, 0)

            y0 = pdfPositionHandling.yo
            x0 = pdfPositionHandling.xo
            text = pdfPositionHandling.text


            from collections import defaultdict

            def list_duplicates(seq):
                tally = defaultdict(list)
                for i, item in enumerate(seq):
                    tally[item].append(i)

                return ((key, locs) for key, locs in tally.items())

            rep = list()
            for each_elem in y0:
                for each_elem2 in y0:
                    if (math.fabs(each_elem - each_elem2) == 1):
                        rep.append((each_elem, each_elem2))

            for t in rep:
                for n, i in enumerate(y0):
                    if i == t[0]:
                        y0[n] = t[1]

            l = []
            for dup in sorted(list_duplicates(y0), reverse=True):
                l.append(dup)

            table_df = pd.DataFrame([])
            res_table = list()
            final_table = list()
            temp_text = ''
            final_table2 = list()




            for dup in sorted(list_duplicates(y0), reverse=True):
                for each_dup in dup[1]:
                    text_append = str(text[each_dup]).replace('\n', '')
                    text_append = text_append
                    res_table.append(text_append)

                final_table.append(res_table)

                while ' ' in res_table:
                    res_table.remove(' ')
                while '  ' in res_table:
                    res_table.remove('  ')
                while '   ' in res_table:
                    res_table.remove('   ')
                while '$' in res_table:
                    res_table.remove('$')

                final_table2.append(res_table)

                res_table = []



            for each_row in final_table:
                table_df = table_df.append(pd.Series(each_row), ignore_index=True)



            s_xo = list(set(x0))
            s_xo = sorted(s_xo)

            for row in final_table2:
                if len(row) == 1:
                    row.clear()



            number_of_clusters = len(max(final_table2, key=len))

            if number_of_clusters<18 and number_of_clusters>15:
                number_of_clusters = 20

            number_of_clusters_list.append(number_of_clusters)
         #   import math
            if (int(math.fabs(number_of_clusters_list[0]-number_of_clusters))==1):
                number_of_clusters = number_of_clusters_list[0]
            #print(number_of_clusters)
            import numpy as np
            kmeans = KMeans(n_clusters=number_of_clusters)
            arr = np.asarray(x0)
            arr = arr.reshape(-1, 1)
            kmeansoutput = kmeans.fit(arr)
            centroids = kmeansoutput.cluster_centers_


            new_centroids = list()
            centroids = centroids.tolist()
            for each_centroid in centroids:
                each_centroid = int(each_centroid[0])
                new_centroids.append(each_centroid)

            new_centroids = sorted(new_centroids)
            new_centroids = sorted(new_centroids)
            #new_centroids = [21, 42, 80, 150, 199, 278, 339, 406, 433,  460, 515, 551]
            #number_of_clusters = number_of_clusters+1



            rep = list()
            for each_elem in y0:
                for each_elem2 in y0:
                    if (math.fabs(each_elem - each_elem2) < 6): #Minimum Distance for new Line
                        rep.append((each_elem, each_elem2))

            for t in rep:
                for n, i in enumerate(y0):
                    if i == t[0]:
                        y0[n] = t[1]

            l2 = list()

            table_df = pd.DataFrame([])
            res_table = list()
            final_table = list()

            for i in range(0, number_of_clusters):
                res_table.append(' ')
                l2.append(' ')

            for dup in sorted(list_duplicates(y0), reverse=True):
                for each_dup in dup[1]:

                    text_append = str(text[each_dup]).replace('\n', '')
                    text_append = text_append.strip()
                    text_append =  re.sub(' +',' ',text_append)
                    cluster = min(range(len(new_centroids)), key=lambda i: abs(new_centroids[i] - x0[each_dup]))

                   # print('clusterr: ', text_append, cluster)

                   # print ('res: ', res_table)
                    leading_sp = len(text_append) - len(text_append.lstrip())
                    if (leading_sp>5):
                        text_append = 'my_pdf_dummy' + '          '+text_append


                    text_append_split = text_append.split('   ')
                    text_append_split_res = []

                    for each_ss in text_append_split:
                        if each_ss!='':
                            each_ss = each_ss.replace('my_pdf_dummy','   ')
                            text_append_split_res.append(each_ss)

                    text_append = text_append.replace('my_pdf_dummy','')
                   # print('tsss: ', text_append_split_res)



                    if (res_table[cluster] != ' ' ):

                      #  print ('tt: ', text_append)
                       # print ('tt: ', cluster)
                        app = str(res_table[cluster] + text_append)


                        res_table[cluster] = app

                    #elif(len(text_append_split_res)>1 and res_table[cluster] != ' '):



                    elif(len(text_append_split_res) > 1):
                        ap = cluster
                        for each_ss in text_append_split_res:

                            try:

                                res_table[ap]=each_ss
                                ap = ap+1
                            except:
                                res_table.insert(ap,each_ss)
                                ap = ap + 1


                    else:

                        res_table[cluster]=text_append
                        #res_table.insert(cluster, text_append)

                for i in range(0, number_of_clusters):
                    res_table.append(' ')

                if not all(' ' == s or s.isspace() for s in res_table):
                    final_table.append(res_table)
                res_table = []
                for i in range(0, number_of_clusters):
                    res_table.append(' ')

            for each_row in final_table:
                table_df = table_df.append(pd.Series(each_row), ignore_index=True)


            all_tables.append(table_df)




    for page_number in range(0,total_pages):
        # print (page_number,"d")
        import PyPDF2# to write contents of pdf to new pdf page by page

        pfr = PyPDF2.PdfFileReader(open(example_file, "rb"))
        orientation = pfr.getPage(0).get('/Rotate')
        # print(orientation,"ori")
        try:
            pfr.decrypt('')
        except:
            pass

        if orientation==180 or orientation==270 or orientation==90:
            # print("in if")
            pdf_in = open(example_file, 'rb')
            pdf_reader = PyPDF2.PdfFileReader(pdf_in)
            pdf_writer = PyPDF2.PdfFileWriter()

            for pagenum in range(pdf_reader.numPages):

                page = pdf_reader.getPage(pagenum)
                # print(pagenum)

                page.rotateClockwise(360-orientation)
                pdf_writer.addPage(page)

            pdf_out = open('rotated5.pdf', 'wb')
            pdf_writer.write(pdf_out)
            pdf_out.close()
            pdf_in.close()

            pfr = PyPDF2.PdfFileReader(open("rotated5.pdf", "rb"))

            pg9 = pfr.getPage(page_number) #extract pg 8
            writer = PyPDF2.PdfFileWriter() #create PdfFileWriter object
            #add pages
            writer.addPage(pg9)
            NewPDFfilename = "input_pdf.pdf"
            with open(NewPDFfilename, "wb") as outputStream:  # create new PDF
                writer.write(outputStream)


        else:
            # print("in else")
            pg9 = pfr.getPage(page_number)  # extract pg 8
            writer = PyPDF2.PdfFileWriter()  # create PdfFileWriter object
            # add pages
            writer.addPage(pg9)
            NewPDFfilename = "input_pdf.pdf"
            with open(NewPDFfilename, "wb") as outputStream:  # create new PDF
                writer.write(outputStream)




        def extract_layout_by_page(pdf_path):#to get layouts of each page in pdf

            # print(pdf_path,"path")
            # print("hello")

            laparams = LAParams()

            fp = open(pdf_path, 'rb')
            parser = PDFParser(fp)
            document = PDFDocument(parser)

            if not document.is_extractable:
                raise PDFTextExtractionNotAllowed

            rsrcmgr = PDFResourceManager()
            device = PDFPageAggregator(rsrcmgr, laparams=laparams)
            interpreter = PDFPageInterpreter(rsrcmgr, device)

            layouts = []
            for page in PDFPage.create_pages(document):
                interpreter.process_page(page)
                layouts.append(device.get_result())
                # print(layouts,"layout")
            return layouts



        page_layouts = extract_layout_by_page(NewPDFfilename)
        # print(page_layouts[0],"ff")
        TEXT_ELEMENTS = [
            pdfminer.layout.LTTextBox,
            pdfminer.layout.LTTextBoxHorizontal,
            pdfminer.layout.LTTextLine,
            pdfminer.layout.LTTextLineHorizontal
        ]

        def flatten(lst):
            # print("list",lst)
            # print(lst)
            return [subelem for elem in lst for subelem in elem]


        def extract_characters(element):
            # print(element)
            # for i in element:
            #     print(i)
            if isinstance(element, pdfminer.layout.LTChar):
                # print("in char")
                # print("1st")

                # print(element)
                # print(element)
                return [element]


            if any(isinstance(element, i) for i in TEXT_ELEMENTS):
                return flatten([extract_characters(e) for e in element])

            # print(element)
            #
            if isinstance(element, list):
                # print(isinstance(element, list))

                return flatten([extract_characters(l) for l in element])

            return []


        final_result = list()
        current_page = page_layouts[0]

        #print('PROCESSING PAGE : ', page_number)

        texts = []
        rects = []

        for e in current_page:
            # print(current_page,"fg")
            # print(e,"ghhh")
            if isinstance(e, pdfminer.layout.LTTextBoxHorizontal):
                texts.append(e)
            elif isinstance(e, pdfminer.layout.LTRect):
                rects.append(e)

        # print(rects,"rects")

        characters = extract_characters(texts)

        import matplotlib.pyplot as plt
        from matplotlib import patches

        def draw_rect_bbox(a, ax, color):
            """
            Draws an unfilled rectable onto ax.
            """
            # print(a[0],"hjjkk")
            ax.add_patch(
                patches.Rectangle(
                    (a[0], a[1]),
                    a[2] - a[0],
                    a[3] - a[1],
                    fill=False,
                    color=color
                )
            )

        def draw_rect(rect, ax, color="black"):
            x0,y0,x1,y1=rect.bbox
            draw_rect_bbox((x0,y0,x1,y1), ax, color)

        xmin, ymin, xmax, ymax = current_page.bbox
        size = 6

        fig, ax = plt.subplots(figsize=(size, size * (ymax / xmax)))

        for rect in rects:
            # print(rect,"hi")
            draw_rect(rect, ax)

        for c in characters:
            draw_rect(c, ax, "red")

        plt.xlim(xmin, xmax)
        plt.ylim(ymin, ymax)
        # plt.show()







        # print(characters)
        # print(rects)


        xmin, ymin, xmax, ymax = current_page.bbox
        # print(xmin, ymin, xmax, ymax)
        size = 6





        def width(rect):
            x0, y0, x1, y1 = rect.bbox
            # print(( x0, y0, x1, y1))
            # print(min( x1 - x0, y1 - y0))
            return min(x1 - x0, y1 - y0)





        def area(rect):
            x0, y0, x1, y1 = rect.bbox

            # print((x0, y0, x1, y1))
            # print((x1 - x0) * (y1 - y0))
            return (x1 - x0) * (y1 - y0)


        # for r in rects:
        #     if width(r)<2:
        #         print(r)






        def cast_as_line(rect):


            x0, y0, x1, y1 = rect.bbox
            # print( x0, y0, x1, y1)

            if x1 - x0 > y1 - y0:
                return (x0, y0, x1, y0, "H")
            else:
                return (x0, y0, x0, y1, "V")


        lines = [cast_as_line(r) for r in rects
                 if width(r) < 2 and
                 area(r) > 1]

        # print(lines)#identify horizontal and vertical lines

        xmin, ymin, xmax, ymax = current_page.bbox
        # print( xmin, ymin, xmax, ymax)
        size = 6

        def does_it_intersect(x, xmin, xmax): #72.504 769.44 225.764 769.9200000000001
            #77
            return (x <= xmax and x >= xmin)

        def find_bounding_rectangle(x, y, lines):

            # print(lines)
            v_intersects=[]



            # for l in lines:
            #     if l[4]=="V" and does_it_intersect(y, l[1], l[3]):
            #         v_intersects.append(l)




            v_intersects = [l for l in lines
                            if l[4] == "V"
                            and does_it_intersect(y, l[1], l[3])]

            # print(v_intersects,"v0")


            h_intersects = [l for l in lines
                            if l[4] == "H"
                            and does_it_intersect(x, l[0], l[2])]
            # print(h_intersects, "h0")


            if len(v_intersects) < 2 or len(h_intersects) < 2:
                # print("ghjkl")
                return None

            v_left = [v[0] for v in v_intersects
                      if v[0] < x]
            # print(v_left)

            v_right = [v[0] for v in v_intersects
                       if v[0] > x]

            # print(v_right)

            if len(v_left) == 0 or len(v_right) == 0:
                return None

            x0, x1 = max(v_left), min(v_right)

            h_down = [h[1] for h in h_intersects
                      if h[1] < y]

            h_up = [h[1] for h in h_intersects
                    if h[1] > y]

            if len(h_down) == 0 or len(h_up) == 0:
                return None

            y0, y1 = max(h_down), min(h_up)

            return (x0, y0, x1, y1)



        from collections import defaultdict
        import math

        box_char_dict = {}

        for c in characters:

            bboxes = defaultdict(int)
            l_x, l_y = c.bbox[0], c.bbox[1]
            bbox_l = find_bounding_rectangle(l_x, l_y,
                                             lines)





            bboxes[bbox_l] += 1

            c_x, c_y = math.floor((c.bbox[0] + c.bbox[2]) / 2), math.floor((c.bbox[1] + c.bbox[3]) / 2)
            bbox_c = find_bounding_rectangle(c_x, c_y, lines)
            bboxes[bbox_c] += 1

            u_x, u_y = c.bbox[2], c.bbox[3]
            bbox_u = find_bounding_rectangle(u_x, u_y, lines)

            bboxes[bbox_u] += 1
            if max(bboxes.values()) == 1:

                bbox = bbox_c
            else:

                bbox = max(bboxes.items(), key=lambda x: x[1])[0]

            if bbox is None:
                continue

            if bbox in box_char_dict.keys():
                box_char_dict[bbox].append(c)
                continue

            box_char_dict[bbox] = [c]

        for x in range(int(xmin), int(xmax), 10):
            for y in range(int(ymin), int(ymax), 10):
                bbox = find_bounding_rectangle(x, y, lines)

                if bbox is None:
                    continue

                if bbox in box_char_dict.keys():
                    continue

                box_char_dict[bbox] = []

        def chars_to_string(chars):

            if not chars:
                return ""
            rows = sorted(list(set(c.bbox[1] for c in chars)), reverse=True)
            text = ""
            for row in rows:
                sorted_row = sorted([c for c in chars if c.bbox[1] == row], key=lambda c: c.bbox[0])
                text = text+' '+"".join(c.get_text() for c in sorted_row)
            return text


        def boxes_to_table(box_record_dict):

            boxes = box_record_dict.keys()
            rows = sorted(list(set(b[1] for b in boxes)), reverse=True)
            table = []
            for row in rows:
                sorted_row = sorted([b for b in boxes if b[1] == row], key=lambda b: b[0])
                table.append([chars_to_string(box_record_dict[b]) for b in sorted_row])
            return table



        result = boxes_to_table(box_char_dict)
        final_result.extend(result)


        if (final_result):

            table_df = pd.DataFrame(final_result)
            all_tables.append(table_df)
        else:
            table_without_border()








    import numpy as np
    all_table_df = pd.DataFrame([])
    for each_table in all_tables:
        all_table_df = all_table_df.append(each_table,ignore_index=True)
        all_table_df = all_table_df.append(pd.Series([np.nan]), ignore_index=True)
        all_table_df = all_table_df.append(pd.Series([np.nan]), ignore_index=True)
        all_table_df = all_table_df.append(pd.Series([np.nan]), ignore_index=True)
        all_table_df = all_table_df.append(pd.Series([np.nan]), ignore_index=True)
        all_table_df = all_table_df.append(pd.Series([np.nan]), ignore_index=True)
        all_table_df = all_table_df.append(pd.Series([np.nan]), ignore_index=True)


    try:
        all_tables = helper_anomaly(len(all_table_df.columns.values))

    except:
        pass

    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    writer = pd.ExcelWriter(desktop + "\output.xlsx", engine='xlsxwriter')
    all_table_df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()









