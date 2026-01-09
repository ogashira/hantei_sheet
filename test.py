

ROW_TITLE: int = 2       # ページの先頭タイトルの行数
ROW_IN_BLOCK: int = 10   # 1つの品番のデータの行数
BLOCKS_IN_PAGE: int = 4  # 1ページに表示する品番数

def calc_pages(block_count):
    count_data: int = block_count
    syou: int = count_data // BLOCKS_IN_PAGE
    amari: int = count_data % BLOCKS_IN_PAGE
    pages: int = syou
    if amari > 0:
        pages = syou + 1

    return pages

def calc_lastRow(blocks_count):

    pages: int = calc_pages(blocks_count)
    blocks_count_lastPage: int = blocks_count % BLOCKS_IN_PAGE  # 余り
    if blocks_count_lastPage == 0:
        blocks_count_lastPage = 4

    beforeLastPage_lastRow: int = ((ROW_TITLE + BLOCKS_IN_PAGE 
                                    * ROW_IN_BLOCK 
                                + BLOCKS_IN_PAGE -1) * (pages-1)) 
    lastPage_lastRow: int =  (ROW_TITLE + blocks_count_lastPage 
                            * ROW_IN_BLOCK + blocks_count_lastPage -1)
    if blocks_count_lastPage == 0:
        lastPage_lastRow = 0
    lastRow: int = beforeLastPage_lastRow + lastPage_lastRow

    return lastRow


    


