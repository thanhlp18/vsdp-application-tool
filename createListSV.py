from ast import Index
from numpy import NaN
import pandas as pd
import toronado
from openpyxl import load_workbook
from zmq import NULL

def isNaN(string):
    return string != string

def returnRealValue(value):
    return value if not(isNaN(value)) else "" 

# Render Student HTML template from a row
def render(student_row):
    html1 = f'''
    <html lang="en"> 
    <head> 
    <meta charset="UTF-8"> 
    <meta http-equiv="X-UA-Compatible" content="IE=edge"> 
    <meta name="viewport" content="width=device-width, initial-scale=1.0"> 
    <title id="title">{returnRealValue(student_row['fullName'])}</title> 
    <link rel="preconnect" href="https://fonts.googleapis.com"> 
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin="crossorigin"> 
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,100;0,200;0,300;0,400;0,500;0,600;0,700;0,800;0,900;1,100;1,200;1,300;1,400;1,500;1,600;1,700;1,800;1,900&amp;display=swap" rel="stylesheet"> 
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css" integrity="sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T" crossorigin="anonymous"> 
    </head> '''
    css = """
        <style>
            html {
                font-family: "Montserrat", sans-serif;
            }

            .logo {
                display: block;
                margin: 24px auto;
                width: 160px;
            }

            .title {
                text-align: center;
                font-size: 3rem;
                color: #14bdd0;
            }

            .content {
                background: #f0feff;
                border-radius: 1.5rem;
                margin-top: 1rem;
            }

            .profile {
                width: 360px;
                height: 360px;
                object-fit: cover;
            }

            .short-text {
                color: #3b3b3b;
                font-weight: bold;
            }

            p {
                margin-bottom: 0.75rem !important;
            }

            .rotatable-image {
                height: 300px;
                width: 300px;
                margin-bottom: 3rem;
                position: relative;
            }

            .rotatable-image img {
                width: 100%;
                height: 100%;
                object-fit: contain;html2 = f'''
            }

            .rotate-btn {
                position: absolute;
                z-index: 999;
                bottom: 0;
                left: 0;
            }

            .modal-content {
                background: rgba(0, 0, 0, 0);
                border: none;
                padding-top: 2rem;
            }
        </style> 
    """
    html2 = f'''
        <body> 
        <div class="container"> 
        <img src="../images/Viethope_Logo.png" alt="" class="logo"> 
        <h1 class="title font-weight-bold">ĐƠN ỨNG TUYỂN VSDP 2022</h1> 
        <section class="information"> 
            <h2 class="mt-5 mb-3 font-weight-bold">1. Thông tin cá nhân</h2> 
            <div class="card conteHọ nt p-4"> 
            <div class="row flex-start mb-3"> 
            <div class="col-2"> 
            <p class="font-weight-bold h4 mb-3">Họ và tên:</p> 
            <p>Tên trường:</p> 
            <p>Ngành học:</p> 
            <p>MSSV:</p> 
            </div> 
            <div class="col-8"> 
            <p class="font-weight-bold h4 mb-3" id="fullName">{returnRealValue(student_row['fullName'])}</p> 
            <p class="short-text" id="university">{returnRealValue(student_row['university'])}</p> 
            <p class="short-text" id="major">{returnRealValue(student_row['major'])}</p> 
            <p class="short-text" id="studentId">{returnRealValue(student_row['studentId'])}</p> 
            </div> 
            </div> 
            <div class="d-flex"> 
            <img src="{returnRealValue(student_row['avatar'])}" class="rounded d-block img-thumbnail profile mr-4" alt="KHÔNG CÓ" id="avatar"> 
            <div class=""> 
            <p>Giới tính theo khai sinh: <span class="short-text" id="gender">{returnRealValue(student_row['gender'])}</span></p> 
            <p>Ngày sinh: <span class="short-text" id="birthday">{returnRealValue(student_row['birthday'])}</span></p> 
            <p> Địa chỉ thường trú: <span class="short-text" id="addressResidence">{returnRealValue(student_row['addressResidence'])}</span> </p> 
            <p>Tỉnh/Thành phố: <span class="short-text" id="city">{returnRealValue(student_row['city'])}</span></p> 
            <p> Địa chỉ liên lạc: <span class="short-text" id="addressContact">{returnRealValue(student_row['addressContact'])}</span> </p> 
            <p>SĐT cá nhân: <span class="short-text" id="phone">{returnRealValue(student_row['phone'])}</span></p> 
            <p>SĐT người thân: <span class="short-text" id="familyPhone">{returnRealValue(student_row['familyPhone'])}</span></p> 
            <p>Số CMND/CCCD: <span class="short-text" id="nationId">{returnRealValue(student_row['nationId'])}</span></p> 
            <p>Email: <span class="short-text" id="email">{returnRealValue(student_row['email'])}</span></p> 
            </div> 
            </div> 
            </div> 
            <div class="card content p-4"> 
            <div class="row"> 
            <div class="col-6"> 
            <p> Sinh viên có đang đi làm thêm: <span class="short-text" id="partTime">{returnRealValue(student_row['partTime'])}</span> </p> 
            </div> 
            <div class="col-6"> 
            <p>Công việc làm thêm: <span class="short-text" id="partTimeJob">{returnRealValue(student_row['partTimeJob'])}</span></p> 
            </div> 
            <div class="col-6"> 
            <p> Thu nhập trung bình mỗi tháng từ công việc làm thêm: <span class="short-text" id="partTimeSalary">{returnRealValue(student_row['partTimeSalary'])}</span> </p> 
            </div> 
            </div> 
            </div> 
        </section> 
        <section class="learning"> 
            <h2 class="mt-5 mb-3 font-weight-bold">2. Thông tin học tập</h2> 
            <div class="card content p-4"> 
            <div class="row"> 
            <div class="col-6"> 
            <p class="font-weight-bold">Trúng tuyển theo phương thức:</p> 
            </div> 
            <div class="col-6"> 
            <p id="method" class="font-weight-bold">{returnRealValue(student_row['method'])}</p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-6"> 
            <p>Điểm thi THPT 2022: <span class="short-text" id="nationExamScore">{returnRealValue(student_row['nationExamScore'])}</span></p> 
            </div> 
            <div class="col-6"> 
            <p>Điểm học bạ THPT: <span class="short-text" id="scoreHighSchool">{returnRealValue(student_row['scoreHighSchool'])}</span></p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-6"> 
            <p>Điểm thi đánh giá năng lực: <span class="short-text" id="scoreCompetency">{returnRealValue(student_row['scoreCompetency'])}</span></p> 
            </div> 
            <div class="col-6"> 
            <p>Tổng điểm các môn trong tổ hợp xét tuyển: <span class="font-weight-bold" id="totalScoreGraduation">{returnRealValue(student_row['totalScoreGraduation'])}</span></p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-6"> 
            <p>Điểm xét tuyển của phương thức đặc cách tốt nghiệp: <span class="short-text" id="totalScoreExceptional">{returnRealValue(student_row['totalScoreExceptional'])}</span></p> 
            </div> 
            </div> 
            </div> 
            <h5 class="text-center mb-3 mt-4" id="notification">Giấy báo nhập học</h5>
            <div class="row justify-content-around">
            <div class="card rotatable-image">
            <img id="imageSrc" class="rounded d-block test-rotate" alt="KHÔNG CÓ" data-toggle="modal" data-target=".image-modal" src="{returnRealValue(student_row['imageSrc'])}"> 
            <div class="rotate-btn card"> 
            <img src="../images/rotate-btn.png"> 
            </div>
            </div>
            </div> 
            <div class="card content p-4"> 
            <div class="row justify-content-around"> 
            <div class="col-4"> 
            <p>Điểm TB lớp 10: <span class="short-text" id="gpa10">{returnRealValue(student_row['gpa10'])}</span></p> 
            </div> 
            <div class="col-4"> 
            <p>Điểm TB lớp 11: <span class="short-text" id="gpa11">{returnRealValue(student_row['gpa11'])}</span></p> 
            </div> 
            <div class="col-4"> 
            <p>Điểm TB lớp 12: <span class="short-text" id="gpa12">{returnRealValue(student_row['gpa12'])}</span></p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-12"> 
            <p> Liệt kê các giải thưởng liên quan đến HỌC TẬP từ THPT: <span class="short-text" id="award"><br>{returnRealValue(student_row['award'])}</span> </p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-12"> 
            <p> Giải thưởng học tập cao nhất: <span class="short-text" id="highestAcademicAward">{returnRealValue(student_row['highestAcademicAward'])}</span> </p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-12"> 
            <p> Liệt kê những học bổng sinh viên đã nhận được tính từ THPT đến nay: <span class="short-text" id="scholarship">{returnRealValue(student_row['scholarship'])}</span> </p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-12"> 
            <p> Sinh viên có thư giới thiệu không: <span class="short-text" id="teacherLetter">{returnRealValue(student_row['teacherLetter'])}</span> </p> 
            </div> 
            </div> 
            </div> 
            '''
    htmlCertificateHighestAcademicAward = f'''
            <h5 class="text-center mb-3 mt-4" id="certificateHighestAcademicAward">Giải thưởng học tập cao nhất</h5>
            <div class="row justify-content-around">
            <div class="card rotatable-image">
            <img id="imageSrc" class="rounded d-block test-rotate" alt="KHÔNG CÓ" data-toggle="modal" data-target=".image-modal" src="{returnRealValue(student_row['certificateHighestAcademicAward'])}"> 
            <div class="rotate-btn card"> 
            <img src="../images/rotate-btn.png"> 
            </div>
            </div>
            </div>
            ''' 
    htmlTeacherLetterFile = f'''
            <h5 class="text-center mb-3 mt-4" id="teacherLetterFile">Thư giới thiệu sinh viên</h5>
            <div class="row justify-content-around">
            <div class="card rotatable-image">
            <img id="imageSrc" class="rounded d-block test-rotate" alt="KHÔNG CÓ" data-toggle="modal" data-target=".image-modal" src="{returnRealValue(student_row['teacherLetterFile'])}"> 
            <div class="rotate-btn card"> 
            <img src="../images/rotate-btn.png"> 
            </div>
            </div>
            </div> 
            '''
    html3 = f'''
        </section> 
        <section class="farmily"> 
            <h2 class="mt-5 mb-3 font-weight-bold">3. Thông tin gia đình</h2> 
            <div class="card content p-4"> 
            <div class="mb-4"> 
            <div class="row"> 
            <div class="col-12"> 
                <p class=""> Họ tên cha: <span class="font-weight-bold" id="fatherName">{returnRealValue(student_row['fatherName'])}</span> </p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-6"> 
                <p> Tình trạng sức khỏe: <span class="short-text" id="fatherHealthStatus">{returnRealValue(student_row['fatherHealthStatus'])}</span> </p> 
            </div> 
            <div class="col-6"> 
                <p>Năm sinh: <span class="short-text" id="fatherBirth">{returnRealValue(student_row['fatherBirth'])}</span></p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-12"> 
                <p> Bệnh (nếu có) hoặc lí do mất sức lao động: <span class="short-text" id="fatherDetailHealth">{returnRealValue(student_row['fatherDetailHealth'])}</span> </p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-6"> 
                <p>Trình độ học vấn: <span class="short-text" id="fatherAcademicLevel">{returnRealValue(student_row['fatherAcademicLevel'])}</span></p> 
            </div> 
            <div class="col-6"> 
                <p>Nghề nghiệp: <span class="short-text" id="fatherJob">{returnRealValue(student_row['fatherJob'])}</span></p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-12"> 
                <p> Thu nhập bình quân hàng tháng: <span class="short-text" id="fatherSalary">{returnRealValue(student_row['fatherSalary'])}</span> </p> 
            </div> 
            </div> 
            </div> 
            <div class="mb-4"> 
            <div class="row"> 
            <div class="col-12"> 
                <p class=""> Họ tên mẹ: <span class="font-weight-bold" id="motherName">{returnRealValue(student_row['motherName'])}</span> </p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-6"> 
                <p> Tình trạng sức khỏe: <span class="short-text" id="motherHealthStatus">{returnRealValue(student_row['motherHealthStatus'])}</span> </p> 
            </div> 
            <div class="col-6"> 
                <p>Năm sinh: <span class="short-text" id="motherBirth">{returnRealValue(student_row['motherBirth'])}</span></p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-12"> 
                <p> Bệnh (nếu có) hoặc lí do mất sức lao động: <span class="short-text" id="motherDetailHealth">{returnRealValue(student_row['motherDetailHealth'])}</span> </p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-6"> 
                <p>Trình độ học vấn: <span class="short-text" id="motherAcademicLevel">{returnRealValue(student_row['motherAcademicLevel'])}</span></p> 
            </div> 
            <div class="col-6"> 
                <p>Nghề nghiệp: <span class="short-text" id="motherJob">{returnRealValue(student_row['motherJob'])}</span></p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-12"> 
                <p> Thu nhập bình quân hàng tháng: <span class="short-text" id="motherSalary">{returnRealValue(student_row['motherSalary'])}</span> </p> 
            </div> 
            </div> 
            </div> 
            </div> 
            <div class="card content p-4"> 
            <div class="row"> 
            <div class="col-12"> 
            <p> Nơi ở của cha mẹ hiện nay: <span class="short-text" id="parentAddress">{returnRealValue(student_row['parentAddress'])}</span> </p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-12"> 
            <p> Tình trạng hôn nhân của cha mẹ: <span class="short-text" id="parentMaritalStatus">{returnRealValue(student_row['parentMaritalStatus'])}</span> </p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-12"> 
            <p> Nếu cha mẹ ly dị thì người còn lại có đóng góp/trợ cấp nuôi con không? <span class="short-text" id="parentSupport">{returnRealValue(student_row['parentSupport'])}</span> </p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-12"> 
            <p> Giấy Chứng Nhận Hộ Nghèo/ Hộ Cận nghèo hoặc Giấy xác nhận gia cảnh khó khăn: <span class="short-text" id="poorHousehold">{returnRealValue(student_row['poorHousehold'])}</span> </p> 
            </div> 
            </div> 
            </div>  
            <div class="card content p-4"> 
            <p class="font-weight-bold">Trả lời chi tiết về thành viên trong gia đình: </p> 
            <div class="col-11"> <span id="memberDetail">{returnRealValue(student_row['memberDetail'])}</span> 
            </div> 
            <p></p> 
            <div class="row"> 
            <div class="col-11"> 
            <p> (1) Số ông/bà đang sống chung </p> 
            </div> 
            <div class="col-1"> 
            <p id="numberGrandparent">{returnRealValue(student_row['numberGrandparent'])}</p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-11"> 
            <p> (2) Số anh/chị/em ĐÃ lập gia đình </p> 
            </div> 
            <div class="col-1"> 
            <p id="numberSiblingMarried">{returnRealValue(student_row['numberSiblingMarried'])}</p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-11"> 
            <p> (3)  Số anh/chị/em CHƯA lập gia đình và đang là sinh viên (không tính sinh viên đang làm đơn): </p> 
            </div> 
            <div class="col-1"> 
            <p id="numberSiblingStudent">{returnRealValue(student_row['numberSiblingStudent'])}</p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-11"> 
            <p> (4) Số anh/chị/em CHƯA lập gia đình, và đang học phổ thông hoặc học nghề </p> 
            </div> 
            <div class="col-1"> 
            <p id="numberSiblingHighSchoolOrApprentice">{returnRealValue(student_row['numberSiblingHighSchoolOrApprentice'])}</p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-11"> 
            <p> (5) Số anh/chị/em CHƯA lập gia đình, và đang làm nông, làm mướn, thất nghiệp, nội trợ HOẶC còn nhỏ chưa đi học: </p> 
            </div> 
            <div class="col-1"> 
            <p id="numberSiblingHighSchoolUnemployed">{returnRealValue(student_row['numberSiblingHighSchoolUnemployed'])}</p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-11"> 
            <p> (6) Số anh/chị/em CHƯA lập gia đình, và làm các nghề khác như công nhân giáo viên, bộ đội, kỹ sư,…. </p> 
            </div> 
            <div class="col-1"> 
            <p id="numberSiblingHighSchoolHaveJob">{returnRealValue(student_row['numberSiblingHighSchoolHaveJob'])}</p> 
            </div> 
            </div> 
            <div class="row"> 
            <div class="col-11"> 
            <p class="font-weight-bold" style="font-size: 1.2rem"> (7) Tổng số thành viên ĐÃ NÊU Ở TRÊN trong gia đình </p> 
            </div> 
            <div class="col-1"> 
            <p class="font-weight-bold" style="font-size: 1.2rem" id="totalNumberMember">{returnRealValue(student_row['totalNumberMember'])}</p> 
            </div> 
            </div> 
            </div> 
        </section> 
        <section class="essay"> 
            <h2 class="mt-5 mb-3 font-weight-bold">4. Bài luận</h2> 
            <div class="card content p-4"> 
            <p class="font-weight-bold"> Câu 1: Bạn hãy trình bày một cách chi tiết nhất về hoàn cảnh gia đình mình hiện nay và những khó khăn, thử thách đã gặp phải khi theo đuổi việc học. (Từ 400 - 1000 từ)</p> 
            </div> 
            <div class="ml-2 mr-2 mt-4 mb-5 text-justify" id="answer1"> 
            <p>{returnRealValue(student_row['answer1'])}</p>
            </div> 
            <div class="card content p-4"> 
            <p class="font-weight-bold">Câu 2: Theo bạn những phẩm chất/ thế mạnh gì đã giúp bạn vượt qua những khó khăn trong quá khứ để tiếp tục việc học? Trong những năm học đại học sắp tới, những thách thức lớn nhất của bạn là gì? Bạn dự định sẽ sử dụng những thế mạnh của mình để vượt qua những thách thức đó như thế nào để theo đuổi con đường học vấn?  Khuyến khích miêu tả chi tiết và có sử dụng ví dụ cụ thể cho những ý trình bày.</p>
            </div> 
            <div class="ml-2 mr-2 mt-4 mb-5 text-justify" id="answer2"> 
            <p>{returnRealValue(student_row['answer2'])}</p>
            </div> 
        </section> 
        <div class="modal fade image-modal" tabindex="-1" role="dialog"> 
            <div class="modal-dialog modal-lg modal-dialog-centered"> 
            <div class="modal-content"> 
            <img id="preview-image" src="" alt="preview"> 
            </div> 
            </div> 
        </div> 
        </div> 
        <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script> ` 
        <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script> 
        <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>  
        </body>
        </html>
    '''
    script = """
    <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script> ` 
  <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script> 
  <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>  
  <script>
    document.querySelectorAll(".rotate-btn").forEach((n) => {
        n.addEventListener("click", (e) => {
            let deg = Number(
                n.parentElement
                    .querySelector("img")
                    .style.transform.replace(/\D+/g, "")
            );
            console.log((deg + 90) % 360);
            n.parentElement.querySelector("img").style.transform = `rotate(${
                (deg + 90) % 360
            }deg)`;
        });
    });

    document.querySelectorAll(".rotatable-image img").forEach((n) => {
        n.addEventListener("click", (e) => {
            let deg = Number(e.target.style.transform.replace(/\D+/g, ""));
            let previewImage = document.querySelector("#preview-image");
            previewImage.setAttribute("src", e.target.getAttribute("src"));
            previewImage.style.transform = `rotate(${deg}deg)`;
        });
    });
</script>  
    """
    with open('./result/%s_%s.html' %(student_row['id'],student_row['fullName']), 'w', encoding='utf-8') as f: 
        f.write(html1)
    with open('./result/%s_%s.html' %(student_row['id'],student_row['fullName']), 'a', encoding='utf-8') as f: 
        f.write(css)
    with open('./result/%s_%s.html' %(student_row['id'],student_row['fullName']), 'a', encoding='utf-8') as f: 
        f.write(html2)
    if(not(isNaN(student_row['certificateHighestAcademicAward']))): 
        with open('./result/%s_%s.html' %(student_row['id'],student_row['fullName']), 'a', encoding='utf-8') as f: 
            f.write(htmlCertificateHighestAcademicAward)
    if(not(isNaN(student_row['teacherLetterFile']))): 
        with open('./result/%s_%s.html' %(student_row['id'],student_row['fullName']), 'a', encoding='utf-8') as f: 
            f.write(htmlTeacherLetterFile)
    with open('./result/%s_%s.html' %(student_row['id'],student_row['fullName']), 'a', encoding='utf-8') as f: 
        f.write(html3)
    with open('./result/%s_%s.html' %(student_row['id'],student_row['fullName']), 'a', encoding='utf-8') as f: 
        f.write(script)
    
# Load student data to dataframe
workbook = pd.ExcelFile('data.xlsx')
df = workbook.parse('Sheet1')
df.replace('nan',0)
# Get nums of row
nums_row = len(df.index)

# for index_row in range(0, nums_row):
#     df.iloc[index_row].fillna("fasdf")
#     for s in df.iloc[index_row]:
#         if(isNaN(s)): print(s)

# Loop to access all colomn in a row
for index_row in range(0, nums_row):
    render(df.iloc[index_row])
