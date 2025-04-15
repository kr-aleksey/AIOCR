import base64
import tomllib

from openai import APIConnectionError, OpenAI, RateLimitError

from errors import OCRError, OCRRateLimitError


def image_to_base64(img: bytes) -> str:
    return base64.b64encode(img).decode('utf-8')


class OpenAIOCR:
    prompt = ('Recognize the text in the image. '
              'As a result, output the text without translation or comments.')


    def __init__(self, model='gpt-4o'):
        with open('settings.toml', 'rb') as f:
            settings = tomllib.load(f)
        self.client = OpenAI(api_key=settings['OpenAI']['api_key'])
        self.model = model

    def _prepare_user_input(self, image: bytes):
        return [
            {
                'role': 'user',
                'content': [
                    {
                        'type': 'input_text',
                        'text': self.prompt
                    },
                    {
                        'type': 'input_image',
                        'image_url': f'data:image/jpeg;base64,{image_to_base64(image)}',
                        'detail': 'auto'
                    },
                ],
            }
        ]

    def ocr(self, image: bytes) -> str:
        user_input = self._prepare_user_input(image)
        try:
            response = self.client.responses.create(model=self.model, input=user_input)
        except APIConnectionError:
            raise OCRError('Нет соединения с сервисом OpenAI. Проверьте интернет.')
        except RateLimitError:
            raise OCRRateLimitError('Превышен лимит запросов в минуту.')
        return response.output_text
