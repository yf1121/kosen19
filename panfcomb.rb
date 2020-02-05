#encoding: Windows-31J
#�f�ڏ��̃f�[�^���`�F�b�N���A��������v���O����
###���\�b�h�̒�`###
# �S�p�������J�E���g
# �����F������
# �Ԃ�l�F�S�p�����̐�
def count_multi_byte(string)
  string.each_char.map{|c| c.bytesize == 1 ? 0 : 1}.reduce(0, &:+) unless string.ascii_only?
end

###���\�b�h�̒�`###
# �S�p��3,���p��2�Ƃ����A������̒����𓱏o
# �����F������
# �Ԃ�l�F����
def count_char(string)
  if string == nil
    haba = 0
  else
    jisu = string.length  #����������
    zenkaku = count_multi_byte(string)  #�S�p�����̃J�E���g
    if zenkaku == nil
      zenkaku = 0
    end
    hankaku = jisu - zenkaku  #���p�����̃J�E���g
    haba = zenkaku + hankaku * 2 / 3.0
  end
  return haba.ceil  #������Ԃ�
end

###���\�b�h�̒�`###
# ���������̉\�������邩����
# �����F������, id, �w�p�Љ��
# �Ԃ�l�F�G���[�R�[�h
def emoji_opt(text, id, lab)
  out = ""
  if text.index("�E") == 0 || text.rindex("�E") == text.length
    disp = "*90%"
    out = "E04a"
  elsif text.include?("�E")
    r = text.split("�E")
    k = 1
    while k < r.size
      str = r[k].slice(0,2)
      if str.scan(/[\p{Han}\p{Katakana}]+/).uniq.length == 0 || r[k].length == 0
        disp = "*70%"
        out = "E04b"
        break
      end
      k += 1
    end
    if text.index("�E") == text.rindex("�E") && (text.index("�E") - text.length / 2) > -(text.length / 6)
      disp = "*40%"
      out = "E04c"
    elsif k == r.size
      disp = "*10%"
      out = "E04d"
    end
  end
  if out.length > 0
    if lab == true
      print("    ", out, "-s: ���������^�f(�w�p�Љ)", disp,"(id=", id, ")\n")
      out = ";" + out + "-s"
    else
      print("    ", out, ": ���������^�f(��`��)", disp,"(id=", id, ")\n")
      out = ";" + out
    end
  end
  return out
end

###���\�b�h�̒�`###
# �S�p��3,���p��2�Ƃ����A������̒����𓱏o
# �����F�z��
# �Ԃ�l�F6�̃^�O�A�C�R��
def tag_judge(array)
  tag6 = Array.new
  tag19 = ["�~�j�̔��R�[�i�[", "�~�j�̌��R�[�i�[", "�~�j�W���R�[�i�[",
      "�q�����y���߂�", "�x�e�ł���", "���ЂƂ肳�܊��}", "�B�e�֎~",
      "�t���b�V���B�e�֎~", "�r�f�I�B�e�E�^���֎~", "�W����i�̐ڐG�֎~",
      "������SNS���e�֎~", "���H�֎~", "�e�C�N�A�E�g��", "�r�����ޏ��",
      "���ԓ���ւ���", "�������z�z����", "�撅��", "�L��", "�J���p��"]
  tagname = ["11_�~�j�̔�", "12_�~�j�̌�", "13_�~�j�W��",
      "14_�q�����y���߂�e�q", "15_�x�e��", "16_���ЂƂ�", "17_�B�e�֎~",
      "18_�t���b�V���B�e�֎~�J�����t��", "19_�r�f�I�B�e�^���֎~", "20_�ڐG�֎~",
      "21_SNS�֎~", "22_���H�֎~", "23_�e�C�N�A�E�g", "24_�r�����ޏ��",
      "25_���ԓ���", "26_������", "27_�撅��", "28_�L��", "29_�J���p"]
  if array != nil
    count = 0
    while count < 19
      if tag6.size > 6
        break
      elsif array.include?(tag19[count])
        tag6.push(tagname[count] + ".ai")
      end
      count += 1
    end
  end
  return tag6[0], tag6[1], tag6[2], tag6[3], tag6[4], tag6[5]  #�Ȃ���nil�ŕԂ�
end

###��������###
starting = Process.clock_gettime(Process::CLOCK_MONOTONIC)  #�������Ԃ̌v���J�n

id = Array.new  #���z��̍쐬
promo = Array.new  #��`���z��쐬
intro = Array.new  #�w�p�Љ�z��쐬
categ = Array.new  #�J�e�S���z��쐬
tag = Array.new  #�^�O�z��쐬
icon = Array.new  #�I���W�i���A�C�R���z��쐬

# ���J�e�S���t�@�C����ǂݍ��݁A�K�v������z��
index = 0  #�J�e�S���z��ϐ���������
Dir.glob("�I�t�B�V�����p���t���b�g�f�ڊ��J�e�S��*.csv").each do |item|  #�t�@�C�����ꗗ�̎擾
  name = File.open(item)
  l_count = 0
  while true  #�������[�v
    line = name.gets  #1�s�ǂݍ���
    if line == nil  #�s���Ȃ���Δ�����
      break
    end
    line.chomp!  #���s����菜��
    tuple = line.gsub(/\"/, '').split(',')  #1�^�v�����Ƃ̃f�[�^�̔z��
    if l_count == 0
      length = tuple.size  #���̕\�̗v�f�����擾
    elsif tuple[0] == ""
      break
    else
      categ[index] = Array.new
      categ[index][0] = tuple[length]  #id���J�e�S���z��[0]�ɕۑ�
      categ[index][1] = tuple[1]  #�啪�ނ��J�e�S���z��[1]�ɕۑ�
      if tuple[1].include?("�Q���^���") || tuple[1].include?("�X�e�[�W")  #�啪��
        tuple_data = tuple[1] + "�]" + tuple[2]
        categ[index][2] = tuple_data  #�����ނ��J�e�S���z��[2]�ɕۑ�
      elsif tuple[1].include?("���i�̔�")  #�啪��
        categ[index][2] = tuple[1] + "�]" + tuple[3]  #�����ނ��J�e�S���z��[2]�ɕۑ�
      elsif tuple[1].include?("�W��")  #�啪��
        categ[index][2] = tuple[1] + "�]" + tuple[4]  #�����ނ��J�e�S���z��[2]�ɕۑ�
      elsif tuple[1].include?("�H�i�̔�")  #�啪��
        categ[index][2] = tuple[1] + "�]" + tuple[5]  #�����ނ��J�e�S���z��[2]�ɕۑ�
      elsif tuple[1].include?("�p�t�H�[�}���X")  #�啪��
        categ[index][2] = tuple[1] + "�]" + tuple[6]  #�����ނ��J�e�S���z��[2]�ɕۑ�
      elsif tuple[1].include?("���̑�")  #�啪��
        categ[index][2] = tuple[1] + "�]" + tuple[7]  #�����ނ��J�e�S���z��[2]�ɕۑ�
      end
      index += 1  #���̊��̏��͏I��
    end
    l_count += 1
  end  #�������[�v�����܂�
end

# ��`���t�@�C����ǂݍ��݁A�K�v������z��
index = 0
Dir.glob("�I�t�B�V�����p���t���b�g��`��*.csv").each do |item|  #�t�@�C�����ꗗ�̎擾
  name = File.open(item)
  l_count = 0
  while true  #�������[�v
    line = name.gets  #1�s�ǂݍ���
    if line == nil  #�s���Ȃ���Δ�����
      break
    end
    line.chomp!  #���s����菜��
    line.chop!  #�����́u"�v����菜��
    tuple = line.split('","')  #1�^�v�����Ƃ̃f�[�^�̔z��
    if l_count == 0
      length = tuple.size  #���̕\�̗v�f�����擾
      tuple[0].slice!("\"")  #�\��id�ۑ�
      regist_id = tuple[0].to_i
      if regist_id == 15
        break
        break
      end
    elsif tuple[0] == "\""
      break
    else
      promo[index] = Array.new
      promo[index][0] = tuple[length]  #id���J�e�S���z��[0]�ɕۑ�
      if length == 4  #�w�p���̏ꍇ
        if tuple[1] == "�w�p���ł���"
          promo[index][1] = tuple[3]  #��`�����`���z��[1]�ɕۑ�
        else
          promo[index][1] = tuple[2]  #��`�����`���z��[1]�ɕۑ�
        end
      elsif length == 2
        promo[index][1] = tuple[1]  #��`�����`���z��[1]�ɕۑ�
      end
      index += 1  #���̊��̏��͏I��
    end
    l_count += 1
  end  #�������[�v�����܂�
end

# �Љ�t�@�C����ǂݍ��݁A�K�v������z��
index = 0
Dir.glob("�I�t�B�V�����p���t���b�g��`��*.csv").each do |item|  #�t�@�C�����ꗗ�̎擾
  name = File.open(item)
  l_count = 0
  while true  #�������[�v
    line = name.gets  #1�s�ǂݍ���
    if line == nil  #�s���Ȃ���Δ�����
      break
    end
    line.chomp!  #���s����菜��
    line.chop!  #�����́u"�v����菜��
    tuple = line.split('","')  #1�^�v�����Ƃ̃f�[�^�̔z��
    if l_count == 0
      length = tuple.size  #���̕\�̗v�f�����擾
      tuple[0].slice!("\"")  #�\��id�ۑ�
      regist_id = tuple[0].to_i
      if regist_id != 15  #�\��id��15�i�w�p�Љ�j�ȊO�̏ꍇ
        break
        break
      end
    elsif tuple[0] == "\""
      break
    else
      intro[index] = Array.new
      intro[index][0] = tuple[length]  #�Љ���w�p�Љ�z��[0]�ɕۑ�
      intro[index][1] = tuple[1]  #�Љ���w�p�Љ�z��[1]�ɕۑ�
      index += 1  #���̊��̏��͏I��
    end
    l_count += 1
  end  #�������[�v�����܂�
end

# �^�O�A�C�R���t�@�C����ǂݍ��݁A�K�v������z��
index = 0
Dir.glob("�I�t�B�V�����p���t���b�g�^�O�A�C�R��.csv").each do |item|  #�t�@�C�����ꗗ�̎擾
  name = File.open(item)
  l_count = 0
  while true  #�������[�v
    line = name.gets  #1�s�ǂݍ���
    if line == nil  #�s���Ȃ���Δ�����
      break
    end
    line.chomp!
    line.chop!
    tuple = line.split('","')  #1�^�v�����Ƃ̃f�[�^�̔z��
    if l_count == 0
      length = tuple.size  #���̕\�̗v�f�����擾
    elsif tuple[0] == "\""
      next
    else
      tag[index] = Array.new
      tag[index][0] = tuple[length]  #id���^�O�z��[0]�ɕۑ�
      tagicon = tuple[1].split(",")
      tag[index] = tag[index] + tag_judge(tagicon)  #�^�O�A�C�R�����^�O�z��ɕۑ�
      index += 1  #���̊��̏��͏I��
    end
    l_count += 1
  end  #�������[�v�����܂�
end
# �I���W�i���A�C�R���t�@�C����ǂݍ��݁A�K�v������z��
index = 0
Dir.glob("�I�t�B�V�����p���t���b�g�w�p�I���W�i���A�C�R��*.csv").each do |item|  #�t�@�C�����ꗗ�̎擾
  name = File.open(item)
  l_count = 0
  while true  #�������[�v
    line = name.gets  #1�s�ǂݍ���
    if line == nil  #�s���Ȃ���Δ�����
      break
    end
    line.chomp!
    line.chop!
    tuple = line.split('","')  #1�^�v�����Ƃ̃f�[�^�̔z��
    if l_count == 0
      length = tuple.size  #���̕\�̗v�f�����擾
    elsif tuple[0] == "\""
      next
    else
      icon[index] = Array.new
      icon[index][0] = tuple[length]  #id���A�C�R���z��[0]�ɕۑ�
      icon[index][1] = tuple[1]  #�t�@�C�������A�C�R���z��[1]�ɕۑ�
      index += 1  #���̊��̏��͏I��
    end
    l_count += 1
  end  #�������[�v�����܂�
end

# �S��惊�X�g��ǂݍ��݁A�K�v������z��
index = 0
Dir.glob("�S���.csv").each do |item|  #�t�@�C�����ꗗ�̎擾
  name = File.open(item)
  l_count = 0
  while true  #�������[�v
    line = name.gets  #1�s�ǂݍ���
    if line == nil  #�s���Ȃ���Δ�����
      break
    end
    line.slice!("\"")
    tuple = line.split('","')  #1�^�v�����Ƃ̃f�[�^�̔z��
    if l_count == 0
    elsif tuple[1] == nil
      next
    elsif tuple[5].include?("\n")
      id[index] = Array.new
      id[index][0] = tuple[0]  #id�����z��[0]�ɕۑ�
      id[index][1] = tuple[1]  #���c�̖������z��[1]�ɕۑ�
      id[index][2] = tuple[3]  #��於�����z��[2]�ɕۑ�
      id[index][4] = tuple[2]  #���c�̖����~�����z��[4]�ɕۑ�
      id[index][5] = tuple[4]  #��於���~�����z��[5]�ɕۑ�
      check = tuple[5]
      id[index][8] = tuple[5].chomp  #���T�v�����z��[8]�ɕۑ�
      while true
        if check.include?("\n")
          line = name.gets
          line.slice!("\"")
          tuple = line.split(',"')  #1�^�v�����Ƃ̃f�[�^�̔z��
          check = tuple[0]
          id[index][8] += tuple[0].chomp
        else
          break
        end
      end
      id[index][3] = tuple[1].chop!  #�Q���敪�����z��[3]�ɕۑ�
      id[index][6] = tuple[2].chop!  #�w�p�g�������z��[6]�ɕۑ�
      id[index][7] = tuple[3].chop!  #�|�p�g�������z��[7]�ɕۑ�
      index += 1  #���̊��̏��͏I��
    else
      id[index] = Array.new
      id[index][0] = tuple[0]  #id�����z��[0]�ɕۑ�
      id[index][1] = tuple[3]  #��於�����z��[1]�ɕۑ�
      id[index][2] = tuple[1]  #���c�̖������z��[2]�ɕۑ�
      id[index][3] = tuple[6]  #�Q���敪�����z��[3]�ɕۑ�
      id[index][4] = tuple[4]  #��於���~�����z��[4]�ɕۑ�
      id[index][5] = tuple[2]  #���c�̖����~�����z��[5]�ɕۑ�
      id[index][6] = tuple[7]  #�w�p�g�������z��[6]�ɕۑ�
      id[index][7] = tuple[8]  #�|�p�g�������z��[7]�ɕۑ�
      id[index][8] = tuple[5]  #���T�v�����z��[8]�ɕۑ�
      index += 1  #���̊��̏��͏I��
    end
    l_count += 1
  end  #�������[�v�����܂�
end

data = Array.new  #�ŏI���`�z��
###�z��f�[�^�̌�����������###
index = 0
while index < id.size  #��搔�����J��Ԃ�
  data[index] = Array.new
  num = 0
  while num < 8  #�S�������i�[
    data[index][num] = id[index][num]
    num += 1
  end
  if data[index][3] == "�X�e�[�W" && data[index][6] == "�Z"
    data[index][8] = "�X�e�[�W"
    data[index][9] = "�X�e�[�W�]�w�p"
  else
    count = 0
    while count < categ.size  #���J�e�S���̏��i�[
      if categ[count][0] == data[index][0]  #id�̈�v
        data[index][8] = categ[count][1]  #�啪��
        data[index][9] = categ[count][2]  #������
        break
      elsif count == categ.size - 1  #�\���Ȃ�
        data[index][8] = ""  #�啪��
        data[index][9] = ""  #������
      end
      count += 1
    end
  end
  count = 0
  while count < promo.size  #��`���̏��i�[
    if promo[count][0] == data[index][0]  #id�̈�v
      data[index][10] = promo[count][1]  #��`��
      break
    elsif count == promo.size - 1  #�\���Ȃ�
      data[index][10] = ""  #��`��
    end
    count += 1
  end
  count = 0
  while count < tag.size  #�^�O�A�C�R��
    if tag[count][0] == data[index][0]
      num = 1
      while num < 7
        data[index][num + 10] = tag[count][num]
        num += 1
      end
    end
    count += 1
  end
  data[index][17] = id[index][8]  #���T�v
  count = 0
  while count < intro.size  #�w�p���Љ�̏��i�[
    if intro[count][0] == data[index][0]  #id�̈�v
      data[index][18] = intro[count][1]  #�Љ
      break
    elsif count == intro.size - 1  #�\���Ȃ�
      data[index][18] = ""  #�Љ
    end
    count += 1
  end
  count = 0
  while count < icon.size  #�w�p�I���W�i���A�C�R���̏��i�[
    if icon[count][0] == data[index][0]  #id�̈�v
      data[index][19] = icon[count][1]  #�I���W�i���A�C�R��
      break
    elsif count == intro.size - 1  #�\���Ȃ�
      data[index][19] = ""  #�I���W�i���A�C�R��
    end
    count += 1
  end
  index += 1
end
###�z��f�[�^�̌��������܂�###

###�f�[�^�`�F�b�N��������###
print("---ERROR����------\n\n")
index = 0
while index < data.size
  data[index][20] = "N"
  #* �����������`�F�b�N *#
  mojisu = count_char(data[index][10])  #��`���̕������J�E���g
  if data[index][6] == "�Z" #�w�p�g�̏ꍇ
    if mojisu > 25
      print("    E01: ��`������*25+", mojisu-25, "��(id=", data[index][0], ")\n")
      data[index][20] += ";E01"  #��`�����߁i�w�p�j�G���[�R�[�hE01
    end
    shokai = count_char(data[index][18])  #�Љ�̕������J�E���g
    if shokai > 105
      print("    E03: �Љ����*105+", shokai-105, "��(id=", data[index][0], ")\n")
      data[index][20] += ";E03"  #�Љ���߁i�w�p�j�G���[�R�[�hE03
    end
    data[index][20] += emoji_opt(data[index][18], data[index][0], true)  #���������`�F�b�N
  else  #��w�p�g�̏ꍇ
    if mojisu > 30
      print("    E02: ��`������*30+", mojisu-30, "��(id=", data[index][0], ")\n")
      data[index][20] += ";E02"  #��`�����߃G���[�R�[�hE02
    end
  end
  #*���������`�F�b�N*#
  data[index][20] += emoji_opt(data[index][10], data[index][0], false)
  index += 1  #���̊��̏��͏I��
end
###�f�[�^�`�F�b�N�����܂�###

# csv�`����panf-out.csv�ɏo��
File.open("panf-out.csv", "w") do |out|
  # csv�`����1�s�ڏo��
  out.print("id,\"��於\",\"���c�̖�\",\"�Q���敪\",\"��於���~\",")
  out.print("\"���c�̖����~\",\"�w�p�g\",\"�|�p�g\",\"�啪��\",\"������\",\"��`��\",")
  out.print("\"tag1\",\"tag2\",\"tag3\",\"tag4\",\"tag5\",\"tag6\",\"���T�v\",")
  out.print("\"�w�p�Љ�\",\"�w�p�I���W�i���A�C�R��\",\"ERROR\"\n")
  # csv�`���ɂQ�s�ڈȍ~�o��
  index = 0  #�J��Ԃ��ϐ�������
  while index < data.size  #�J��Ԃ�
    out.print(data[index][0].to_i)
    count = 1  #�J��Ԃ��ϐ��̏�����
    while count < data[index].size  #�S�v�f�����o��
      out.print(",\"", data[index][count], "\"")
      count += 1  #�J��Ԃ��ϐ��X�V
    end
    out.print("\n")  #��؂�̉��s
    index += 1  #�J��Ԃ��ϐ��X�V
  end
end

print("\n\n��{���𓝍����Apanf-out.csv�ɏ����o���܂����B\n")
print("�����������{�ꏊ���𓝍����܂�...\n\n")


#�z�u���𓝍����A�y�[�W������U��v���O����
###���\�b�h��`###
#�G���A�ԍ���t�^���郁�\�b�h
# �����F�G���A��
# �Ԃ�l�F�G���A�ԍ�
def areaindex(areaname, tent)
  areaname = areaname.to_s
  tent = tent.to_s
  if areaname.include?("�����r����")
    if tent.include?("C") || tent.include?("D")
      areano = "23"
    elsif tent.include?("E") || tent.include?("F")
      areano = "25"
    else
      areano = "25"
    end
  elsif areaname.include?("�}���َ���")
    if tent.include?("C") || tent.include?("D")
      areano = "03"
    elsif tent.include?("E")
      areano = "05"
    elsif tent.include?("F")
      areano = "07"
    else
      areano = "07"
    end
  elsif areaname.include?("��w��َ���")
    areano = "31"
  elsif areaname.include?("5C������")
    if tent.include?("A") || tent.include?("B")
      areano = "35"
    else
      areano = "37"
    end
  elsif areaname.include?("2A���E3A������")
    areano = "01"
  elsif areaname.include?("1A���E1C���E1D������")
    if tent.include?("A") || tent.include?("B")
      areano = "19"
    else
      areano = "21"
    end
  else
    areano = "43"  #�G���A����/�s��
  end
  return areano
end

#�G���A�ԍ���t�^���郁�\�b�h
# �����F����
# �Ԃ�l�F�G���A�ԍ�
def areaindex_okunai(area, room)
  area = area.to_s
  if area.include?("��R�G���A")
    areano = "09"
  elsif area.include?("��Q�G���A")
    areano = "13"
  elsif area.include?("��P�G���A")
    areano = "27"
  elsif area.include?("��كG���A")
    areano = "33"
  elsif area.include?("��")
    if /5[A-Z][4-7]+/ =~ room
      areano = "39"
    elsif /5[A-Z]+/ =~ room
      areano = "40"
    elsif /6[A-Z][4-7]+/ =~ room
      areano = "41"
    else
      areano = "42"
    end
  elsif area.include?("���̑�")
    if room.include?("�����}����")
      areano = "09"
    elsif room.include?("��������")
      areano = "09"
    elsif room.include?("�J�w�L�O")
      areano = "42"
    else
      areano = "43"
    end
  else
    areano = "44"  #�G���A����/�s��/�e��
  end
  return areano
end

#�e���g�\�[�g�R�[�h��t�^���郁�\�b�h
# �����F�e���g�ԍ�
# �Ԃ�l�F�\�[�g�p�R�[�h, �e���g�ԍ�
def tentindex(tent)
  str = tent.to_s
  chr = str.slice!(0)
  code = chr.to_s.tr("A-I", "1-9")
  str = sprintf("%02d", str.to_i).to_s
  code = code + str
  return code.to_s
end

#��敪�ރA�C�R���̃t�@�C�������擾���郁�\�b�h
# �����F������
# �Ԃ�l�F��敪�ރA�C�R���̃t�@�C����
def kikakuicon_add(category)
  if category == nil
    name = nil
  elsif /^�Q���^���+/ =~ category
    if category.include?("��")  #���k��E���k��̓J�E���^�[�A�C�R��
      name = "counter.ai"
    else  #����ȊO�̎Q���^�͎Q���A�C�R��
      name = "joint.ai"
    end
  elsif /^���i�̔�+/ =~ category
    name = "present.ai"
  elsif /^�W���]�|+/ =~ category
    name = "picture.ai"
  elsif /^�W���]�w+/ =~ category
    name = "pencil.ai"
  elsif /^�W���]+/ =~ category
    name = "projector.ai"
  elsif /^�H�i�̔�+/ =~ category
    if category.include?("�X�C�[�c")  #�X�C�[�c�̓P�[�L�A�C�R��
      name = "cake.ai"
    elsif category.include?("�h�����N") || category.include?("�i��")
      name = "drink.ai"  #�����n�̓R�b�v�A�C�R��
    else  #����ȊO�̐H�i�̔��̓��X�g�����A�C�R��
      name = "restaurant.ai"
    end
  elsif /^�p�t�H�[�}���X+/ =~ category
    name = "microphone.ai"
  elsif /^�X�e�[�W+/ =~ category
    name = "stage.ai"
  else
    name = "etc.ai"
  end
  return name
end

#�w�p���y�[�W�̌f�ڏ�����t�^���郁�\�b�h
# �����Fid, �w�p�n�b�V��
# �Ԃ�l�F�w�p���Q�ƃy�[�W�ԍ�, �w�p���Q�ƔԒn
def tsukulab(id, lab)
  if lab.size > 0 && lab.key?(id.to_s)
    page, address = lab[id.to_s]
  else
    page = (lab.size).div(7) + 1
    address = (lab.size) % 7 + 1
    store_p = [page, address]
    lab.store(id.to_s, store_p)
  end
  return page, address
end

#���[���h�N�C�Y�����[�A�C�R����t�^���郁�\�b�h
# �����Fid, ���[���h�N�C�Y�����[���z��
# �Ԃ�l�F�t�@�C����
def jitsuiadd(kikakuid, wqlist)
  result = ""
  i = 0
  while i < wqlist.size
    if kikakuid.to_s == wqlist[i][0].to_s
      result = "���[���h�N�C�Y�����[.ai"
    end
    i += 1
  end
  return result
end

###��������###
require 'csv'
id = Array.new  #���z��̍쐬
okugai = Array.new  #���O�z��쐬
okunai = Array.new  #�����z��쐬
icon = Array.new  #�A�C�R���z��쐬
ord = Array.new  #��ʔz��쐬
stg = Array.new  #�X�e�[�W�z��쐬
lab = Hash.new  #�w�p�n�b�V���쐬
wqlist = Array.new  #���[���h�N�C�Y�����[���z��쐬

# �����i���O�j��CSV��ǂݍ��݁A�K�v������z��
index = 0
filename = Dir.glob("*�����*���O*.csv").max_by { |fn| File.birthtime(fn) }
if filename == nil
  print("���O�̊�����CSV�t�@�C����������܂���ł����B\n�����t�H���_�Ɋ�����CSV�t�@�C����u���Ă�����s���Ă��������B\n")
  exit
end
CSV.foreach(filename) do |tuple|  #�t�@�C�����ꗗ�̎擾
  if tuple[0] == "id"
    next
  elsif tuple[0] == nil
    break
  elsif tuple[1].to_s.include?("(��掫��)")
    next
  else
    okugai[index] = Array.new
    okugai[index][0] = tuple[0].to_i  #id�����O�z��[0]�ɕۑ�
    okugai[index][1] = tuple[4]  #�G���A�����O�z��[1]�ɕۑ�
    if tuple[5].to_s.include?("/")
      multiarea = tuple[4].split("/")
      multitent = tuple[5].split("/")
      k = 0
      while k < multiarea.size
        if k != 0
          okugai[index + k] = Array.new
          okugai[index + k][0] = tuple[0].to_i  #id�����O�z��[0]�ɕۑ�
          okugai[index + k][1] = tuple[4]  #�G���A�����O�z��[1]�ɕۑ�
        end
        okugai[index + k][2] = areaindex(multiarea[k], multitent[k])  #�G���A�̏��ԕt��
        okugai[index + k][4] = sprintf("%3s", multitent[k]).strip  #�e���g�ԍ������O�z��[4]�ɕۑ�
        okugai[index + k][3] = tentindex(multitent[k])  #�e���gID�����O�z��[3]�ɕۑ�
        okugai[index + k][5] = tuple[6]  #2�������O�z��[5]�ɕۑ�
        okugai[index + k][6] = tuple[7]  #3�������O�z��[6]�ɕۑ�
        okugai[index + k][7] = tuple[8]  #4�������O�z��[7]�ɕۑ�
        okugai[index + k][8] = tuple[11]  #�e���g�O�������O�z��[8]�ɕۑ�
        okugai[index + k][9] = ";E05"  #�ړ��̃G���[�R�[�h���L�^
        if okugai[index + k][3] == 0  #�e���g�ԍ���l�̏ꍇ
          okugai[index + k][9] += ";E07"  #�e���g�ԍ��s���̃G���[�R�[�h���L�^
        end
        if okugai[index][2] == 95  #�G���A�s���̏ꍇ
          okugai[index][9] += ";E08"  #�G���A�s���̃G���[�R�[�h���L�^
        end
        k += 1
      end
      index = index + k  #���̊��̔z�u���͏I��
    else
      okugai[index][2] = areaindex(tuple[4], tuple[5])  #�G���A�̏��ԕt��
      okugai[index][4] = sprintf("%3s", tuple[5]).strip   #�e���g�ԍ������O�z��[4]�ɕۑ�
      okugai[index][3] = tentindex(tuple[5])  #�e���gID�����O�z��[3]�ɕۑ�
      okugai[index][5] = tuple[6]  #2�������O�z��[5]�ɕۑ�
      okugai[index][6] = tuple[7]  #3�������O�z��[6]�ɕۑ�
      okugai[index][7] = tuple[8]  #4�������O�z��[7]�ɕۑ�
      okugai[index][8] = tuple[11]  #�e���g�O�������O�z��[8]�ɕۑ�
      okugai[index][9] = ""  #ERROR�Ȃ�
      if okugai[index][3] == 0  #�e���g�ԍ���l�̏ꍇ
        okugai[index][9] += ";E07"  #�e���g�ԍ��s���̃G���[�R�[�h���L�^
      end
      if okugai[index][2] == 95  #�G���A�s���̏ꍇ
        okugai[index][9] += ";E08"  #�G���A�s���̃G���[�R�[�h���L�^
      end
      index += 1  #���̊��̔z�u���͏I��
    end
  end
end  #CSV�ǂݍ��݂����܂�

# �����i�����j��CSV��ǂݍ��݁A�K�v������z��
index = 0
filename = Dir.glob("*�����*����*.csv").max_by { |fn| File.birthtime(fn) }
if filename == nil
  print("�����̊�����CSV�t�@�C����������܂���ł����B\n�����t�H���_�Ɋ�����CSV�t�@�C����u���Ă�����s���Ă��������B\n")
  exit
end
CSV.foreach(filename) do |item|  #�t�@�C�����ꗗ�̎擾
  if item[0] == "id"
    next
   elsif item[0] == nil
    break
  elsif item[2].to_s.include?("��掫��") || item[2].to_s.include?("��撆�~") || item[2].to_s.include?("��掩��")
    next
  else
    okunai[index] = Array.new
    okunai[index][0] = item[0].to_i  #id�������z��[0]�ɕۑ�
    okunai[index][1] = item[2]  #�G���A�������z��[1]�ɕۑ�
    if item[3].to_s.include?(",")  #���������̏ꍇ
      if item[3].to_s.include?(":")  #���ɂ����ꏊ�ύX�̏ꍇ
        multiroom = item[3].split(",")
        k = 0
        while k < multiroom.size
          if k != 0
            okunai[index + k] = Array.new
            okunai[index + k][0] = item[0].to_i  #id�������z��[0]�ɕۑ�
            okunai[index + k][1] = item[2]  #�G���A�������z��[1]�ɕۑ�
          end
          okunai[index + k][2] = areaindex_okunai(item[2], multiroom[k].sub(/...:/, "").strip)  #�G���A�̏��ԕt��
          okunai[index + k][3] = multiroom[k].sub(/...:/, "").strip  #�����������z��[3]�ɕۑ�
          okunai[index + k][4] = multiroom[k].sub(/...:/, "").strip  #�����������z��[4]�ɕۑ�
          okunai[index + k][5] = "�~"  #2�������O�z��[5]�ɕۑ�
          okunai[index + k][6] = "�~"  #3�������O�z��[6]�ɕۑ�
          okunai[index + k][7] = "�~"  #4�������O�z��[7]�ɕۑ�
          if /�O���+/ =~ multiroom[k]
            okunai[index + k][5] = item[4]  #2�������O�z��[5]�ɕۑ�
          end
          if /1����+/ =~ multiroom[k]
            okunai[index + k][6] = item[5]  #3�������O�z��[6]�ɕۑ�
          end
          if /2����+/ =~ multiroom[k]
            okunai[index + k][7] = item[6]  #4�������O�z��[7]�ɕۑ�
          end
          okunai[index + k][8] = item[7]  #�e���g�O�������O�z��[8]�ɕۑ�
          okunai[index + k][9] = ";E05"  #�ړ��̃G���[�R�[�h���L�^
          if okugai[index][2] == 99  #�G���A�s���̏ꍇ
            okugai[index][9] += ";E08"  #�G���A�s���̃G���[�R�[�h���L�^
          end
          k += 1
        end
        index = index + k  #���̊��̔z�u���͏I��
      else
        room_m = item[3].split(",")
        item[3] = room_m[0]
        j = 1
        while j < room_m.size
          if /\d[A-Z]+/ =~ room_m[j].to_s
            item[3] = item[3] + "/" + room_m[j]
          else
            item[3] = item[3] + "," + room_m[j]
          end
          j += 1
        end
        multiroom = item[3].split("/")
        k = 0
        while k < multiroom.size
          if k != 0
            okunai[index + k] = Array.new
            okunai[index + k][0] = item[0].to_i  #id�������z��[0]�ɕۑ�
            okunai[index + k][1] = item[2]  #�G���A�������z��[1]�ɕۑ�
          end
          okunai[index + k][2] = areaindex_okunai(item[2], multiroom[k])  #�G���A�̏��ԕt��
          okunai[index + k][3] = multiroom[k]  #�����������z��[3]�ɕۑ�
          okunai[index + k][4] = item[3]  #�����������z��[4]�ɕۑ�
          okunai[index + k][5] = item[4]  #2�������O�z��[5]�ɕۑ�
          okunai[index + k][6] = item[5]  #3�������O�z��[6]�ɕۑ�
          okunai[index + k][7] = item[6]  #4�������O�z��[7]�ɕۑ�
          okunai[index + k][8] = item[7]  #�e���g�O�������O�z��[8]�ɕۑ�
          okunai[index + k][9] = ";E06"  #�����ꏊ�̃G���[�R�[�h���L�^
          if okunai[index + k][2] == 99  #�G���A�s���̏ꍇ
            okunai[index + k][9] += ";E08"  #�G���A�s���̃G���[�R�[�h���L�^
          end
          k += 1
        end
        index = index + k  #���̊��̔z�u���͏I��
      end
    else
      okunai[index][2] = areaindex_okunai(item[2], item[3])  #�G���A�̏��ԕt��
      okunai[index][3] = item[3]  #�����������z��[3]�ɕۑ�
      okunai[index][4] = item[3]  #�����������z��[4]�ɕۑ�
      i = 5
      while i < 9
        okunai[index][i] = item[i - 1]  #�����z��ɕۑ�
        i += 1
      end
      okunai[index][9] = ""  #ERROR�Ȃ�
      if okunai[index][2] == 99  #�G���A�s���̏ꍇ
        okunai[index][9] += ";E08"  #�G���A�s���̃G���[�R�[�h���L�^
      end
      index += 1  #���̊��̔z�u���͏I��
    end
  end
end  #CSV�ǂݍ��݂����܂�

# �X�e�[�W�^�C���e�[�u����CSV��ǂݍ��݁A�K�v������z��
index = 0
Dir.glob("*TT*.csv").each do |filename|
  if filename == nil
    print("�X�e�[�W�^�C���e�[�u����CSV�t�@�C����������܂���ł����B\n�����t�H���_�ɃX�e�[�W�^�C���e�[�u����CSV�t�@�C����u���Ă�����s���Ă��������B\n")
    exit
  end
  date = 2  #���̏����l
  place = ""  #�ꏊ�̏����l
  CSV.foreach(filename) do |tuple|  #�t�@�C�����ꗗ�̎擾
    if tuple[0] != nil && tuple[1] == nil
      name = filename.split("_")
      File.open("panf-note.txt", "w") do |note|
        note.print("�X�e�[�W:", name[0], ",", date, "��,���e:", tuple[0], "\n")
      end
      print("\n���L��panf-note.txt�ɏ����o���܂����B\n\n")
    elsif tuple[0].to_s.include?("�O���")
      date = 2
    elsif tuple[0].to_s.include?("1����")
      date = 3
      if tuple[0].to_s.include?("�z�[��")
        place = "�i�z�[���j"
      elsif tuple[0].to_s.include?("�u��")
        place = "�i�u���j"
      end
    elsif tuple[0].to_s.include?("2����")
      date = 4
      if tuple[0].to_s.include?("�z�[��")
        place = "�i�z�[���j"
      elsif tuple[0].to_s.include?("�u��")
        place = "�i�u���j"
      end
    elsif tuple[0] == nil
      next
    else
      name = filename.split("_")
      stg[index] = Array.new
      stg[index][0] = tuple[0].to_s  #��於���X�e�[�W�z��[0]�ɕۑ�
      stg[index][1] = date  #�����X�e�[�W�z��[1]�ɕۑ�
      stg[index][2] = name[0] + place  #�ꏊ���X�e�[�W�z��[2]�ɕۑ�
      stg[index][3] = tuple[1]  #�J�n�������X�e�[�W�z��[3]�ɕۑ�
      stg[index][4] = tuple[2]  #�I���������X�e�[�W�z��[4]�ɕۑ�
      stg[index][5] = sprintf("%02d", index)  #�������X�e�[�W�z��[6]�ɕۑ�
      stg[index][6] = tuple[3]  #id���X�e�[�W�z��[6]�ɕۑ�
      index += 1  #���̊��̔z�u���͏I��
    end
  end  #CSV�ǂݍ��݂����܂�
end

# ���[���h�N�C�Y�����[��惊�X�g��ǂݍ��݁A�K�v������z��
index = 0
Dir.glob("*���[���h�N�C�Y�����[*.txt").each do |item|  #�t�@�C�����ꗗ�̎擾
  CSV.foreach(item,:col_sep => "\t") do |row|
    wqlist[index] = Array.new
    wqlist[index] = row
    index += 1
  end
end

# panf-out.csv��ǂݍ��݁A�K�v������z��
index = 0
Dir.glob("panf-out*.csv").each do |item|  #�t�@�C�����ꗗ�̎擾
  name = File.open(item)
  line = name.gets  #1�s�ǂݍ���
  while true  #�������[�v
    line = name.gets  #1�s�ǂݍ���
    if line == nil  #�s���Ȃ���Δ�����
      break
    end
    line.chomp!
    tuple = line.split(',"')  #1�^�v�����Ƃ̃f�[�^�̔z��
    if /\A\(���\s+/ =~ tuple[1]  #��掫�ނ��Ƃ΂�
      next
    else
      id[index] = Array.new
      id[index][0] = tuple[0].to_i  #id�����z��[0]�ɕۑ�
      i = 1
      while i < 21
        id[index][i] = tuple[i].chop  #���z��ɕۑ�
        i += 1
      end
      if id[index][10] == ""
        id[index][20] += ";E10"  #��`����l�̃G���[�R�[�h���L�^
      end
      if id[index][9] == ""
        id[index][20] += ";E09"  #���J�e�S����l�̃G���[�R�[�h���L�^
      end
      if id[index][6] == "�Z"
        if id[index][18] == ""
          id[index][20] += ";E11"  #�Љ��l�̃G���[�R�[�h���L�^
        end
        if id[index][19] == ""
          id[index][20] += ";E12"  #�w�p�I���W�i���A�C�R���t�@�C������l�̃G���[�R�[�h���L�^
        end
      end
      index += 1  #���̊��̏��͏I��
    end
  end  #�������[�v�����܂�
end
if id.size == 0
  print("panf-out.csv��������܂���ł����B\ncomb.rb�����s���āA�����t�H���_��panf-out.csv��u���Ă�����s���Ă��������B\n")
  exit
end

###�y�[�W���Ƀ\�[�g###
ord = okugai + okunai  #��ʊ��z����쐬
#�ꏊ���ɕ��ёւ�(�y�[�W��=>�\�[�g�R�[�h=>�O���=>1����)
ord.sort! {|a, b| (a[2].to_i <=> b[2].to_i).nonzero? || (a[3].to_s <=> b[3].to_s).nonzero? || (b[5].to_s <=> a[5].to_s).nonzero? || (b[6].to_s <=> a[6].to_s)}
stg.sort! {|a, b| (a[1] <=> b[1]).nonzero? || (a[2] <=> b[2]).nonzero? || (a[5] <=> b[5])}

###��ʊ��̔z��f�[�^�̌���###
index = 0
while index < ord.size  #��搔�����J��Ԃ�
  error = ord[index][9].to_s  #�G���[�R�[�h�̋L�^
  count = 0
  while count < id.size
    if ord[index][0] == id[count][0]  #����
      num = 1
      while num < 21
        ord[index][num + 8] = id[count][num]
        num += 1
      end
      break
    end
    count += 1
  end
  if count == id.size  #��������񂪌�����Ȃ������ꍇ
    ord[index][9] = ""
    ord[index][28] = error + ";E33"  #�G���[�R�[�h��Ԃ�
  else
    ord[index][28] = ord[index][28].to_s + error  #�G���[�R�[�h�̃f�[�^����
    ord[index][29] = kikakuicon_add(ord[index][17])  #��敪�ރA�C�R���̃t�@�C�������i�[
  end
  if ord[index][14] == "�Z"  #�w�p�g�Ȃ�΁A�����{�Q�ƃy�[�W���i�[
    ord[index][30], ord[index][31] = tsukulab(ord[index][0], lab)
    ord[index][32] = "01_Tsukulab_mini.ai"  #�w�p�g�A�C�R���̃t�@�C�������i�[
  end
  if ord[index][15] == "�Z"  #�|�p�g�Ȃ�΁A�|�p�A�C�R���̃t�@�C�������i�[
    if ord[index][32] == nil
      ord[index][32] = "04_geijutsuwaku.ai"
      ord[index][33] = jitsuiadd(ord[index][0], wqlist)  #�\���w���σA�C�R���̃t�@�C�����i�[
    else
      ord[index][33] = "04_geijutsuwaku.ai"
      ord[index][35] = jitsuiadd(ord[index][0], wqlist)  #�\���w���σA�C�R���̃t�@�C�����i�[
    end
  end
  if ord[index][32] == nil  #�܂������w���σA�C�R�����t�^����Ă��Ȃ��ꍇ
    ord[index][32] = jitsuiadd(ord[index][0], wqlist)  #�\���w���σA�C�R���̃t�@�C�����i�[
  end
  ord[index][34] = sprintf("%03d", index + 1)  #��ʃV���A���i���o�[�̕t�^(��L�[)
  index += 1  #���̊��̏����͏I������
end

###�X�e�[�W���̔z��f�[�^�̌���###
index = 0
while index < stg.size  #��搔�����J��Ԃ�
  count = 0
  while count < id.size
    if id[count][3] == "�X�e�[�W" && (stg[index][0] == id[count][1] || stg[index][6] == id[count][0].to_s)  #����
      stg[index][6] = id[count][20].to_s  #�G���[�R�[�h�̓���
      num = 0
      while num < 20
        stg[index][num + 7] = id[count][num]
        num += 1
      end
      break
    end
    count += 1
  end
  if stg[index][6].to_s.include?(";") == false
    count = 0
    while count < id.size
      if stg[index][0] == id[count][2]  #�c�̖��Ō��������݂�
        stg[index][6] = id[count][20].to_s  #�G���[�R�[�h�̓���
        stg[index][7] = id[count][0]  #id�̓���
        stg[index][8] = id[count][2]  #��於�Ɋ�{���̒c�̖�������
        stg[index][9] = id[count][1]  #���c�̖��Ɋ�{���̊�於������
        num = 3
        while num < 20
          stg[index][num + 7] = id[count][num]
          num += 1
        end
        stg[index][6] = stg[index][6].to_s + ";E31"
        break
      end
      count += 1
    end
    if stg[index][6].to_s.include?(";") == false  #�����R�Â����s�̃G���[�R�[�h��Ԃ�
      stg[index][6] = "N;E30"
    end
  end
  if stg[index][13].to_s == "�Z"  #�w�p�g�Ȃ�΁A�����{�Q�ƃy�[�W���i�[
    stg[index][27], stg[index][28] = tsukulab(stg[index][7], lab)
    stg[index][29] = "01_Tsukulab_mini.ai"  #�w�p�g�A�C�R���̃t�@�C�������i�[
  end
  index += 1
end

###��ʊ��z���csv�`����panf-ord.csv�ɏo��###
File.open("panf-ord.csv", "w") do |out|
  # csv�`����1�s�ڏo��
  out.print("ORD�ر�,id,\"�ꏊ\",\"�y�[�W����\",\"�\�[�g�R�[�h\",\"�e���g�ԍ�/��������\",")
  out.print("\"2��\",\"3��\",\"4��\",\"�e���g�O���\",\"��於\",\"���c�̖�\",")
  out.print("\"�Q���敪\",\"��於���~\",\"���c�̖����~\",\"�w�p�g\",\"�|�p�g\",")
  out.print("\"�啪��\",\"������\",\"��`��\",\"�w����icon1\",\"�w����icon2\",\"�w����icon3\",\"tag1\",\"tag2\",\"tag3\",\"tag4\",")
  out.print("\"tag5\",\"tag6\",\"���T�v\",\"�w�p�Љ�\",\"�w�p�I���W�i���A�C�R��\",\"�w�p�Q�ƃy�[�W\",\"�w�p�Q�ƔԒn\",\"��敪�ރA�C�R��\",\"ERROR\"\n")
  # csv�`���ɂQ�s�ڈȍ~�o��
  index = 0  #�J��Ԃ��ϐ�������
  while index < ord.size  #�J��Ԃ�
    out.print(ord[index][34])  #��ʃV���A���i���o�[���1��ɏo��
    count = 0  #�J��Ԃ��ϐ��̏�����
    while count < 19  #�S�v�f�����o��
      out.print(",\"", ord[index][count], "\"")
      count += 1  #�J��Ԃ��ϐ��X�V
    end
    out.print(",\"", ord[index][32], "\",\"", ord[index][33], "\",\"", ord[index][35], "\"")  #�w���σA�C�R���t�@�C�����̏o��
    while count < 28  #�S�v�f�����o��
      out.print(",\"", ord[index][count], "\"")
      count += 1  #�J��Ԃ��ϐ��X�V
    end
    out.print(",\"", ord[index][30], "\",\"", ord[index][31], "\"")  #�w�p�y�[�W�ԍ��̏o��
    out.print(",\"", ord[index][29], "\"")  #��敪�ރA�C�R�������̏o��
    out.print(",\"", ord[index][28], "\"")  #�G���[�R�[�h�����̏o��
    out.print("\n")  #��؂�̉��s
    index += 1  #�J��Ԃ��ϐ��X�V
  end
end

###�X�e�[�W���z���csv�`����panf-stg.csv�ɏo��###
File.open("panf-stg.csv", "w") do |out|
  # csv�`����1�s�ڏo��
  out.print("STG�ر�,id,\"���t\",\"�ꏊ\",\"�J�n����\",\"�I������\",\"��於\",\"���c�̖�\",")
  out.print("\"�Q���敪\",\"��於���~\",\"���c�̖����~\",\"�w�p�g\",\"�|�p�g\",")
  out.print("\"�啪��\",\"������\",\"��`��\",\"�w����icon\",\"���T�v\",\"�w�p�Љ�\",\"�w�p�I���W�i���A�C�R��\",\"�w�p�Q�ƃy�[�W\",\"�w�p�Q�ƔԒn\",\"ERROR\"\n")
  # csv�`���ɂQ�s�ڈȍ~�o��
  index = 0  #�J��Ԃ��ϐ�������
  while index < stg.size  #�J��Ԃ�
    out.print(stg[index][5])  #�X�e�[�W�V���A���i���o�[���1��ɏo��
    out.print(",\"", stg[index][7], "\"")  #id���2��ɏo��
    count = 1  #�J��Ԃ��ϐ��̏�����
    while count < 5  #�S�v�f�����o��
      out.print(",\"", stg[index][count], "\"")
      count += 1  #�J��Ԃ��ϐ��X�V
    end
    out.print(",\"", stg[index][0], "\"")  #��於���o��
    count = 9  #�J��Ԃ��ϐ��̏�����
    while count < 18  #�S�v�f�����o��
      out.print(",\"", stg[index][count], "\"")
      count += 1  #�J��Ԃ��ϐ��X�V
    end
    out.print(",\"", stg[index][29], "\"")  #�w�p�A�C�R���t�@�C�������o��
    out.print(",\"", stg[index][24], "\"")  #���T�v���o��
    out.print(",\"", stg[index][25], "\"")  #�w�p�Љ���o��
    out.print(",\"", stg[index][26], "\"")  #�I���W�i���A�C�R�����o��
    out.print(",\"", stg[index][27], "\",\"", stg[index][28], "\"")  #�w�p�Q�ƃy�[�W���o��
    out.print(",\"", stg[index][6], "\"")  #�G���[�R�[�h���o��
    out.print("\n")  #��؂�̉��s
    index += 1  #�J��Ԃ��ϐ��X�V
  end
end

ending = Process.clock_gettime(Process::CLOCK_MONOTONIC)  #�������Ԃ̌v���I��
print("\n�����f�[�^��panf-ord.csv��panf-stg.csv�ɏ����o���܂����B\n")
print("�����J�n����̌o�ߎ���: ", ending - starting, "�b\n(Enter�������ďI�����܂�)\n")
ok = gets.chomp!
exit
