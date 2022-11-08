mod easy_excel;

use proc_macro::TokenStream;
use syn::{parse_macro_input, DeriveInput};

#[proc_macro_derive(easy_excel, attributes(excel, excel_group))]
pub fn easy_excel(input: TokenStream) -> TokenStream {
    let input = parse_macro_input!(input as DeriveInput);
    match do_expand(input) {
        Ok(token_stream) => token_stream.into(),
        Err(e) => e.to_compile_error().into(),
    }
    // TokenStream::default()
}

fn do_expand(input: DeriveInput) -> syn::Result<proc_macro2::TokenStream> {
    easy_excel::EasyExcelContext::from(input).render()
}
