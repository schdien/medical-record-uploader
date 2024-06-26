use std::{io::{self, BufReader, Write, Read},error::Error, fs::File};
use reqwest::{Client,header::{HeaderMap, HeaderValue}, Response};
use calamine::{open_workbook, Xlsx, Reader, DataType, Range};
use base64;


enum State{
    User,
    Rows(MedicClient),
    Choose(MedicClient)
}

struct MedicRecord {
    id: String,
    task: String,
    time: String,
    num: String,
    basic_info: String,
    disease: String,
    result: String,
}

impl MedicRecord {
    fn new(row:&[DataType])->MedicRecord {
        let medic_record = MedicRecord{
            id: row[0].to_string(),
            task: row[3].to_string(),
            time: row[4].to_string(),
            num: row[5].to_string(),
            basic_info: row[6].to_string(),
            disease: row[7].to_string(),
            result: base64::encode(row[8].to_string()).replace("/", "#").replace("+", "_").replace("=", "^"),
        };
        medic_record
    }
}


struct MedicClient{
    client: Client,
    headers: HeaderMap,

    url: String,
    //cookie: String,
    view_state: String,
    depart_list: String,
    depart_id: String,
    depart_order: String,
    train_depart: String,
    task_type: String,
    curr_guid: String,
    user_id: String,

    records: Vec<MedicRecord>,

    total_num: usize,
    success_num: usize,
    fail_num: usize,
    fail_ids: Vec<String>,

}

impl MedicClient{
    fn new(
        sheet:Range<DataType>,
        raw_data: &str
    )->MedicClient{
            let mut client = MedicClient{
                client: Client::new(),
                headers: HeaderMap::new(),
                url: String::new(),
                view_state: String::new(),
                depart_list: String::new(),
                depart_id: String::new(),
                depart_order: String::new(),
                train_depart: String::new(),
                task_type: String::new(),
                curr_guid: String::new(),
                user_id: String::new(),

                records: Vec::new(),

                total_num: 0,
                success_num: 0,
                fail_num: 0,
                fail_ids: Vec::new(),
            };
             
            let data:Vec<&str> = raw_data.split_ascii_whitespace().collect();
            
            for i in 0..data.len(){
                match &data[i] as &str {
                    "Accept:" => {client.headers.insert("Accept", HeaderValue::from_str(data[i+1].trim_end()).unwrap());},
                    "Cookie:" => {client.headers.insert("Cookie", HeaderValue::from_str(data[i+1].trim_end()).unwrap());},
                    "Host:" => {client.headers.insert("Host", HeaderValue::from_str(data[i+1].trim_end()).unwrap());},
                    "Origin" => {client.headers.insert("Origin", HeaderValue::from_str(data[i+1].trim_end()).unwrap());},
                    "Referer:" => {
                        client.url.push_str(data[i+1].trim_end());
                        client.headers.insert("Referer", HeaderValue::from_str(data[i+1].trim_end()).unwrap());
                    },
                    "User-Agent:" => {client.headers.insert("User-Agent", HeaderValue::from_str(data[i+1].trim_end()).unwrap());},
                    
                    "__VIEWSTATE:" => client.view_state.push_str(data[i+1].trim_end()),
                    "selDepartmentList:" => client.depart_list.push_str(data[i+1].trim_end()),
                    "selResourceType:" => client.task_type.push_str(data[i+1].trim_end()),
                    "txtCurrGUID:" => client.curr_guid.push_str(data[i+1].trim_end()),
                    "txtUserID:" => client.user_id.push_str(data[i+1].trim_end()),
                    "txtCurrDepartID:" => client.depart_id.push_str(data[i+1].trim_end()),
                    "txtOrder:" => client.depart_order.push_str(data[i+1].trim_end()),
                    "txtUserTrainDepartID:" => client.train_depart.push_str(data[i+1].trim_end()),
                    _ => {},
                }
            }

            for row in sheet.rows(){
                client.records.push(MedicRecord::new(row));
            }

            client.total_num = client.records.len();
            
            return client;
    }


    async fn post_form(&mut self, start:usize, end: usize, raw_task_ids: &[&str]){
        let mut task_ids:Vec<String> = Vec::new();
        for raw_task_id in raw_task_ids{
            let task_id = "chkTrainItem$".to_owned() + *raw_task_id;
            task_ids.push(task_id);
        }
        for record in self.records.iter().skip(start-1).take(end-start+1){
            let mut form = vec![
                ("__EVENTTARGET",""),
                ("__EVENTARGUMENT",""),
                ("__LASTFOCUS",""),
                ("__VIEWSTATE",&self.view_state),
                ("__VIEWSTATEENCRYPTED",""),
                ("selDepartmentList",&self.depart_list),
                ("selResourceType",&self.task_type),
                ("txtOperateTime",&record.time),
                ("txtPatientNumber",&record.num),
                ("txtPatientInfo",&record.basic_info),
                ("txtDisease1",&record.disease),
                //("fileFile0",""),
                ("Button2","提  交"),
                ("txtCurrGUID",&self.curr_guid),
                ("txtCurrPatientID",""),
                ("txtCurrID",""),
                ("txtUserID",&self.user_id),
                ("txtCurrDepartID",&self.depart_id),
                ("txtOrder",&self.depart_order),
                ("txtUserTrainDepartID",&self.train_depart),
                ("txtpdx",""),
                ("txtManagerID",""),
                ("txtty",""),
                ("txtDepartID",""),
                ("hidFileAdress",""),
                ("strTypeN","医技报告"),
                ("strPaAdd",""),
                ("txtPatientId",""),
                ("litHistory",""),
                ("hidUrl",""),
                ("txtType","医技报告"),
                ("strStartTime",""),
                ("strDiagnosis",""),
                ("strTarget",""),
                ("strMeasures",""),
                ("strEvaluate",""),
                ("strEndTime",""),
                ("txtApPatientDiagnosis4E",""),
                ("txtApPatientDiagnosis5E",""),
                ("txtApPatientDiagnosis6E",""),
                ("txtApPatientDiagnosis7E",""),
                ("txtApPatientDiagnosis8E",""),
                ("txtApPatientDiagnosis9E",""),
                ("txtContentE", &record.result),
                ];
            //自动判断任务名称
            if raw_task_ids.len() == 0 {
                match &self.depart_id as &str{
                    "51" => {
                        match &record.task as &str{
                            //51 
                            "膀胱肿瘤" => form.push(("chkTrainItem$0","on")),
                            "胆结石" => form.push(("chkTrainItem$1","on")),
                            "房间隔缺损" => form.push(("chkTrainItem$2","on")),
                            "风心病二尖瓣狭窄" => form.push(("chkTrainItem$3","on")),
                            "肝癌" => form.push(("chkTrainItem$4","on")),
                            "肝血管瘤" => form.push(("chkTrainItem$5","on")),
                            "肝硬化" => form.push(("chkTrainItem$6","on")),
                            "高血压病" => form.push(("chkTrainItem$7","on")),
                            "冠心病" => form.push(("chkTrainItem$8","on")),
                            "卵巢肿瘤" => form.push(("chkTrainItem$9","on")),
                            "乳腺肿瘤" => form.push(("chkTrainItem$10","on")),
                            "肾结石" => form.push(("chkTrainItem$11","on")),
                            "肾肿瘤" => form.push(("chkTrainItem$12","on")),
                            "室间隔缺损" => form.push(("chkTrainItem$13","on")),
                            "心肌病" => form.push(("chkTrainItem$14","on")),
                            _ => panic!("error: unexpected task name: {}", &record.task)
                        }
                    }
                    "91" => {
                        match &record.task as &str{
                            //91
                            "鼻咽癌" => form.push(("chkTrainItem$0","on")),
                            "肠梗阻" => form.push(("chkTrainItem$1","on")),
                            "胆石症" => form.push(("chkTrainItem$2","on")),
                            "肺结核" => form.push(("chkTrainItem$3","on")),
                            "肺脓肿" => form.push(("chkTrainItem$4","on")),
                            "肺心病" => form.push(("chkTrainItem$5","on")),
                            "肺炎" => form.push(("chkTrainItem$6","on")),
                            "肺肿瘤" => form.push(("chkTrainItem$7","on")),
                            //"" => form.push(("chkTrainItem$8","on")),
                            "肝癌" => form.push(("chkTrainItem$9","on")),
                            "肝血管瘤" => form.push(("chkTrainItem$10","on")),
                            "肝硬化" => form.push(("chkTrainItem$11","on")),
                            "高血压性心脏病" => form.push(("chkTrainItem$12","on")),
                            "骨肿瘤" => form.push(("chkTrainItem$13","on")),
                            "甲状腺肿瘤" => form.push(("chkTrainItem$14","on")),
                            "结直肠癌" => form.push(("chkTrainItem$15","on")),
                            "淋巴瘤" => form.push(("chkTrainItem$16","on")),
                            "慢性支气管炎肺气肿" => form.push(("chkTrainItem$17","on")),
                            "脑血管意外" => form.push(("chkTrainItem$18","on")),
                            "乳腺癌" => form.push(("chkTrainItem$19","on")),
                            "软组织肿瘤" => form.push(("chkTrainItem$20","on")),
                            "肾脏肿瘤" => form.push(("chkTrainItem$21","on")),
                            "食管癌" => form.push(("chkTrainItem$22","on")),
                            "食管静脉曲张" => form.push(("chkTrainItem$23","on")),
                            "唾液腺肿瘤" => form.push(("chkTrainItem$24","on")),
                            "胃、十二指肠溃疡" => form.push(("chkTrainItem$25","on")),
                            "胃癌" => form.push(("chkTrainItem$26","on")),
                            "胰腺癌" => form.push(("chkTrainItem$27","on")),
                            "支气管扩张" => form.push(("chkTrainItem$28","on")),
                            "纵隔肿瘤" => form.push(("chkTrainItem$29","on")),
                            _ => panic!("error: unexpected task name: {}", &record.task)
                        }
                    }
                    "188" => {
                        match &record.task as &str{
                            //188 病理诊断
                            "冰冻切片诊断" => form.push(("chkTrainItem$0","on")),
                            "临床病理讨论会并在上级医生指导下完成病例讨论的病理报告" => form.push(("chkTrainItem$5","on")),
                            "特殊染色及免疫组化染色在病理诊断和鉴别诊断中的应用原则和准确判断结果的技能" => form.push(("chkTrainItem$8","on")),
                            "科内病理读片会诊" => form.push(("chkTrainItem$4","on")),
                            _ => panic!("error: unexpected task name: {}", &record.task)
                        }
                    }
                    "187" => {
                        if record.task.contains("免疫组织化学染色原理（抗体）") {
                            form.push(("chkTrainItem$2","on"));
                            form.push(("chkTrainItem$9","on"));
                        }
                        else if record.task.contains("常规苏木素") {
                            form.push(("chkTrainItem$4","on"));
                            form.push(("chkTrainItem$10","on"));
                            form.push(("chkTrainItem$15","on"));
                        }
                        else if &record.task == "免疫组化所造成的人为变化和特异性控制" {
                            form.push(("chkTrainItem$8","on"));
                        }
                        else if &record.task == "电镜制片的基本方法及技巧" {
                            form.push(("chkTrainItem$6","on"));
                        }
                        else {
                            panic!("error: unexpected task name: {}", &record.task);
                        } 
                    }
                    _ => panic!("error: unexpected department id: {}", self.depart_id)
                }
            }   
            //设置任务名称
            else{
                for task_id in task_ids.iter() {
                    form.push((task_id,"on"));              //("chkTrainItem$x","on"),
                }
            }

            //println!("{:#?}",form);
            let res = self.client.post(&self.url).headers(self.headers.clone()).form(&form).send().await.unwrap();
            let text =res.text().await.unwrap();
            //println!("{}",text);
            
            if text.contains("成功保存") {
                self.success_num += 1;
            }
            else {
                self.fail_num += 1;
                self.fail_ids.push(record.id.to_string());
            }
            if self.fail_num ==10 {
                break;
            }
            
            print!("\rUpload: {}/{} - Fail: {} - Current patient: {}",self.success_num,self.total_num,self.fail_num,record.id);
            io::stdout().flush().unwrap(); 
        }
        println!("\n{:#?}",self.fail_ids);
    }

}



#[tokio::main]
async fn main() -> Result<(), Box<dyn Error>>{

    let mut state = State::User;
    loop{
        state = match state {
            State::User => {
                println!("Enter the form path:");
                let mut path = String::new();
                io::stdin().read_line(&mut path).unwrap();
                let path:&str = path.trim_end();

                println!("Enter the user information path:");
                let mut user = String::new();
                io::stdin().read_line(&mut user).unwrap();
                let user:&str = user.trim_end();

                let mut data_file  = File::open(user).unwrap();
                let mut raw_data = String::new();
                data_file.read_to_string(&mut raw_data).unwrap();

                let sheet = open_workbook::<Xlsx<BufReader<File>>,&str>(path).unwrap().worksheet_range("Sheet1").unwrap().unwrap();
                let client = MedicClient::new(sheet,&raw_data);
                State::Rows(client)
            }
            State::Rows(mut client) => {
                println!("Enter: \"start row,end row,training program index\"");
                let mut input = String::new();
                io::stdin().read_line(&mut input).unwrap();
                let mut s = input.trim_end().split(",");
                let start:usize = s.next().unwrap().parse().unwrap();
                let end:usize = s.next().unwrap().parse().unwrap();
                let raw_task_ids:Vec<&str> = s.collect();//3

                client.post_form(start,end,&raw_task_ids).await;
                
                State::Choose(client)
            }
            State::Choose(client) => {
                println!("\n\nA => New user\nB => New Rows\nOthers => Exit");
                let mut choice = String::new();
                io::stdin().read_line(&mut choice).unwrap();
                match choice.trim_end() {
                    "A" | "a" => State::User,
                    "B" | "b" => State::Rows(client),
                    _ => break,
                }
            }
        }
    }

    Ok(())
}


