from xlsxx.coord import (
    index_to_column, column_to_index,
    ref_to_coord, coord_to_ref, split_ref, modify_ref,
    split_range_ref, range_ref_to_coord, modify_range_ref,
    get_coord, get_range_coord
)

def test_column_index():
    assert "A" == index_to_column(0)
    assert "C" == index_to_column(2)
    assert "Z" == index_to_column(25)
    assert "AA" == index_to_column(26)
    assert "AB" == index_to_column(27)
    assert "BA" == index_to_column(52)
    assert "ZZ" == index_to_column(701)
    assert "AAA" == index_to_column(702)

    assert 0 == column_to_index("A")
    assert 26 == column_to_index("AA")
    assert 702 == column_to_index("AAA")


def test_refcoord():
    assert (0,0) == ref_to_coord("A1")
    assert (3,26) == ref_to_coord("AA4")
    
    assert "A1" == coord_to_ref((0, 0))
    assert "AAB26" == coord_to_ref((25, 703))

    assert ("A", "1") == split_ref("A1")
    assert ("AB", "32") == split_ref("AB32")



def test_ref():
    assert "A1" == modify_ref("", row="1", col="A")
    assert "A1" == modify_ref("", row=0, col=0)
    assert "B2" == modify_ref("", row="2", col=1)

    assert "F3" == modify_ref("A3", col=5)
    assert "Z3" == modify_ref("A3", col="Z")
    assert "A6" == modify_ref("A3", row=5)
    assert "A6" == modify_ref("A3", row="6")


def test_range_ref():
    assert ("A2", "B3") == split_range_ref("A2:B3")
    assert ("AB23", "BC34") == split_range_ref("AB23:BC34")

    assert ((1, 0), (2, 1)) == range_ref_to_coord("A2:B3")

    assert "A2:B3" == modify_range_ref("", head="A2", tail="B3")
    assert "A2:B3" == modify_range_ref("A2:A3", tail="B3")
    assert "A2:B3" == modify_range_ref("A1:B1", headrow="2", tailrow="3")
    assert "A2:B3" == modify_range_ref("A1:B1", headrow=1, tailrow=2)


def test_get_coord():
    pass
