// use easy_excel::easy_excel::*;
use easy_excel::*;
// use crate::easy_exc;

#[allow(dead_code)]
#[derive(easy_excel, Debug)]
pub struct User {
    #[excel(width = "100")]
    name: String,
    #[excel(order = 1, title = "姓名", width = "10")]
    opt_name: String,
    #[excel(order = 9, title = "性别")]
    sex: Option<bool>,
    #[excel(order = 2, title = "年龄")]
    age: u8,
    #[excel(order = 4, title = "list")]
    list: Vec<User>,
}

fn main() {
    let co = vec![
        User {
            name: "user1".to_string(),
            opt_name: "opt_name_user1".to_string(),
            sex: Some(true),
            age: 1,
            list: vec![],
        },
        User {
            name: "user2".to_string(),
            opt_name: "opt_name_user2".to_string(),
            sex: None,
            age: 2,
            list: vec![],
        },
    ];
    check_path(std::path::Path::new(
        "/Users/wanyifan/Downloads/baiduyun/test2.xlsx",
    ));

    User::write_excel(
        co,
        std::path::Path::new("/Users/wanyifan/Downloads/baiduyun/test2.xlsx"),
    );
}
fn check_path(path: &std::path::Path) -> bool {
    let p_path = path.parent().unwrap();
    if p_path.exists() {
        !path.exists()
    } else {
        std::fs::create_dir(p_path).unwrap();
        true
    }
}
