// use easy_excel::easy_excel::*;
use easy_excel::*;
// use crate::easy_exc;

#[allow(dead_code)]
#[derive(easy_excel, Debug)]
pub struct User {
    #[excel(index = 1, title = "姓名")]
    name: String,
    #[excel(index = 2, title = "性别")]
    sex: String,
    #[excel(index = 3, title = "年龄")]
    age: u8,
    #[excel(index = 4, title = "list")]
    list: Vec<User>,
}

fn main() {
    let co = vec![
        User {
            name: "user1".to_string(),
            sex: "sex".to_string(),
            age: 1,
            list: vec![],
        },
        User {
            name: "user2".to_string(),
            sex: "sex".to_string(),
            age: 2,
            list: vec![],
        },
    ];
    User::write_excel(
        co,
        std::path::Path::new("/Users/wanyifan/Downloads/baiduyun/test2.xlsx"),
    );
}
